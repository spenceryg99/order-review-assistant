use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use chrono::Local;
use rand::prelude::IndexedRandom;
use regex::Regex;
use reqwest::blocking::{Client, Response};
use reqwest::header::{HeaderMap, HeaderValue, AUTHORIZATION, CONTENT_TYPE, USER_AGENT};
use reqwest::Url;
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};
use walkdir::WalkDir;

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct AppSettings {
    api_base_url: String,
    jst_login_url: String,
    jst_cookie: String,
    jst_owner_co_id: String,
    jst_authorize_co_id: String,
    jst_uid: String,
    ai_api_base: String,
    ai_api_key: String,
    ai_model: String,
    review_prompt_template: String,
    image_root_dir: String,
    output_root_dir: String,
    order_column_name: String,
    images_per_product: usize,
}

impl Default for AppSettings {
    fn default() -> Self {
        Self {
            api_base_url: "http://192.168.1.166/index.php/api/Yingdao/".to_string(),
            jst_login_url: "https://www.erp321.com/".to_string(),
            jst_cookie: String::new(),
            jst_owner_co_id: "14805587".to_string(),
            jst_authorize_co_id: "14805587".to_string(),
            jst_uid: "21419081".to_string(),
            ai_api_base: "https://dashscope.aliyuncs.com/compatible-mode/v1".to_string(),
            ai_api_key: String::new(),
            ai_model: "qwen-plus".to_string(),
            review_prompt_template:
                "请生成1条符合真人评价特点的淘宝商品好评，需要评价的商品是：{product_name}。直接返回评价内容，不要其他任何描述文字。"
                    .to_string(),
            image_root_dir: String::new(),
            output_root_dir: String::new(),
            order_column_name: "订单号".to_string(),
            images_per_product: 5,
        }
    }
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ValidateCookieRequest {
    cookie: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ValidateCookieResult {
    valid: bool,
    message: String,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TestPromptRequest {
    settings: AppSettings,
    product_name: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct TestPromptResult {
    prompt: String,
    review: String,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RunRequest {
    settings: AppSettings,
    excel_path: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct RunResult {
    output_dir: String,
    summary_file: String,
    total_rows: usize,
    total_orders: usize,
    total_products: usize,
    generated_reviews: usize,
    missing_products: Vec<String>,
    failed_items: Vec<String>,
}

#[derive(Debug, Clone)]
struct ImageProfile {
    folder_name: String,
    aliases: Vec<String>,
    images: Vec<PathBuf>,
}

#[derive(Debug)]
struct SummaryItem {
    order_id: String,
    product_name: String,
    matched_folder: String,
    review_file: String,
    image_count: usize,
    status: String,
}

#[derive(Debug)]
struct RowOrder {
    row_number: usize,
    order_id: String,
}

#[derive(Debug, Serialize)]
struct OpenAiChatRequest {
    model: String,
    messages: Vec<OpenAiMessage>,
    temperature: f32,
}

#[derive(Debug, Serialize)]
struct OpenAiMessage {
    role: String,
    content: String,
}

fn settings_path(app: &AppHandle) -> Result<PathBuf> {
    let dir = app
        .path()
        .app_config_dir()
        .context("无法获取应用配置目录")?;
    fs::create_dir_all(&dir).context("创建配置目录失败")?;
    Ok(dir.join("settings.json"))
}

fn normalize_text(input: &str) -> String {
    let punctuation = [
        '，', '。', '！', '？', '；', '：', '“', '”', '‘', '’', '（', '）', '【', '】', '《',
        '》', ',', '.', '!', '?', ';', ':', '"', '\'', '-', '_', '~', '`', ' ', '\t', '\r',
        '\n', '/', '\\', '*',
    ];

    input
        .trim()
        .to_lowercase()
        .chars()
        .filter(|ch| !punctuation.contains(ch))
        .collect()
}

fn sanitize_filename(name: &str) -> String {
    let re = Regex::new(r#"[<>:"/\\|?*\x00-\x1F]"#).expect("regex compile should succeed");
    let cleaned = re.replace_all(name, "_");
    let trimmed = cleaned.trim();
    if trimmed.is_empty() {
        "unnamed".to_string()
    } else {
        trimmed.to_string()
    }
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(v) => v.trim().to_string(),
        Data::Float(v) => {
            if (*v - v.trunc()).abs() < f64::EPSILON {
                format!("{}", *v as i64)
            } else {
                format!("{}", v)
            }
        }
        Data::Int(v) => v.to_string(),
        Data::Bool(v) => v.to_string(),
        Data::DateTime(v) => v.to_string(),
        Data::DateTimeIso(v) => v.to_string(),
        Data::DurationIso(v) => v.to_string(),
        Data::Error(_) => String::new(),
    }
}

fn chunk_orders(order_ids: &[String], chunk_size: usize) -> Vec<Vec<String>> {
    let mut chunks = Vec::new();
    let mut idx = 0;
    while idx < order_ids.len() {
        let end = std::cmp::min(idx + chunk_size, order_ids.len());
        chunks.push(order_ids[idx..end].to_vec());
        idx = end;
    }
    chunks
}

fn build_http_client() -> Result<Client> {
    Client::builder()
        .timeout(std::time::Duration::from_secs(35))
        .build()
        .context("构建 HTTP 客户端失败")
}

fn get_view_state(client: &Client, cookie: &str) -> Result<String> {
    let response = client
        .get("https://www.erp321.com/app/order/order/list.aspx")
        .query(&[("_c", "jst-epaas")])
        .header(
            USER_AGENT,
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        )
        .header("cookie", cookie)
        .send()
        .context("请求聚水潭订单页失败")?;

    ensure_success(response)?.text().context("读取聚水潭页面失败")
        .and_then(|text| {
            let re = Regex::new(r#"id="__VIEWSTATE" value="([^"]+)""#)
                .context("构建 VIEWSTATE 正则失败")?;
            let value = re
                .captures(&text)
                .and_then(|caps| caps.get(1))
                .map(|m| m.as_str().to_string())
                .ok_or_else(|| anyhow!("Cookie 可能已失效：无法解析 __VIEWSTATE"))?;
            Ok(value)
        })
}

fn ensure_success(response: Response) -> Result<Response> {
    let status = response.status();
    if status.is_success() {
        Ok(response)
    } else {
        let body = response.text().unwrap_or_default();
        Err(anyhow!("HTTP {}: {}", status.as_u16(), body))
    }
}

fn query_products_once(
    client: &Client,
    cookie: &str,
    view_state: &str,
    order_ids: &[String],
) -> Result<HashMap<String, Vec<String>>> {
    if order_ids.is_empty() {
        return Ok(HashMap::new());
    }

    let joined = order_ids.join(",");
    let filter = format!(r#"[{{\"k\":\"so_id\",\"v\":\"{}\",\"c\":\"@=\"}}]"#, joined);
    let callback = format!(
        r#"{{"Method":"LoadDataToJSON","Args":["1","{}","{{}}"]}}"#,
        filter
    );

    let timestamp = Local::now().timestamp_millis().to_string();
    let form_data = [
        ("__VIEWSTATE", view_state.to_string()),
        ("__VIEWSTATEGENERATOR", "C8154B07".to_string()),
        ("insurePrice", String::new()),
        ("_jt_page_count_enabled", String::new()),
        ("_jt_page_increament_enabled", "true".to_string()),
        ("_jt_page_increament_page_mode", String::new()),
        ("_jt_page_increament_key_value", String::new()),
        ("_jt_page_increament_business_values", String::new()),
        ("_jt_page_increament_key_name", "o_id".to_string()),
        ("_jt_page_size", "50".to_string()),
        ("_jt_page_action", "1".to_string()),
        ("fe_node_desc", String::new()),
        ("receiver_state", String::new()),
        ("receiver_city", String::new()),
        ("receiver_district", String::new()),
        ("receiver_address", String::new()),
        ("receiver_name", String::new()),
        ("receiver_phone", String::new()),
        ("receiver_mobile", String::new()),
        ("check_name", String::new()),
        ("check_address", String::new()),
        ("fe_remark_type", "single".to_string()),
        ("fe_flag", String::new()),
        ("fe_is_append_remark", String::new()),
        ("feedback", String::new()),
        ("__CALLBACKID", "JTable1".to_string()),
        ("__CALLBACKPARAM", callback),
    ];

    let response = client
        .post("https://www.erp321.com/app/order/order/list.aspx")
        .query(&[
            ("_c", "jst-epaas"),
            ("ts___", timestamp.as_str()),
            ("am___", "LoadDataToJSON"),
        ])
        .header(
            USER_AGENT,
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        )
        .header("x-requested-with", "XMLHttpRequest")
        .header("origin", "https://www.erp321.com")
        .header("referer", "https://www.erp321.com/app/order/order/list.aspx?_c=jst-epaas")
        .header("cookie", cookie)
        .form(&form_data)
        .send()
        .context("查询订单商品失败")?;

    let text = ensure_success(response)?.text().context("读取订单查询响应失败")?;
    let (_, payload) = text
        .split_once('|')
        .ok_or_else(|| anyhow!("订单查询响应格式异常: {}", text))?;

    let first_json: serde_json::Value =
        serde_json::from_str(payload).context("解析订单查询 JSON 失败")?;
    let return_value = first_json["ReturnValue"]
        .as_str()
        .ok_or_else(|| anyhow!("订单查询结果缺少 ReturnValue"))?;
    let second_json: serde_json::Value =
        serde_json::from_str(return_value).context("解析订单查询 ReturnValue 失败")?;

    let mut map = HashMap::new();
    let datas = second_json["datas"]
        .as_array()
        .ok_or_else(|| anyhow!("订单查询结果 datas 为空"))?;

    for row in datas {
        let order_id = row["so_id"].as_str().unwrap_or_default().trim().to_string();
        if order_id.is_empty() {
            continue;
        }

        let products = row["items"]
            .as_array()
            .map(|items| {
                items
                    .iter()
                    .filter_map(|item| item["name"].as_str())
                    .map(|name| name.replace('*', "_").trim().to_string())
                    .filter(|name| !name.is_empty())
                    .collect::<Vec<String>>()
            })
            .unwrap_or_default();

        if !products.is_empty() {
            map.insert(order_id, products);
        }
    }

    Ok(map)
}

fn fetch_products_by_orders(
    client: &Client,
    cookie: &str,
    order_ids: &[String],
) -> Result<HashMap<String, Vec<String>>> {
    let view_state = get_view_state(client, cookie)?;
    let mut all_map = HashMap::new();

    for chunk in chunk_orders(order_ids, 50) {
        let mut chunk_map = query_products_once(client, cookie, &view_state, &chunk)?;

        let chunk_set: HashSet<String> = chunk.iter().cloned().collect();
        let found_set: HashSet<String> = chunk_map.keys().cloned().collect();
        let missing: Vec<String> = chunk_set.difference(&found_set).cloned().collect();

        for order in missing {
            if let Ok(single_map) = query_products_once(client, cookie, &view_state, &[order.clone()]) {
                for (k, v) in single_map {
                    chunk_map.insert(k, v);
                }
            }
        }

        all_map.extend(chunk_map);
    }

    Ok(all_map)
}

fn collect_order_rows(excel_path: &str, order_column_name: &str) -> Result<Vec<RowOrder>> {
    let mut workbook = open_workbook_auto(excel_path)
        .with_context(|| format!("打开 Excel 失败: {}", excel_path))?;

    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| anyhow!("Excel 中没有工作表"))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .context("读取工作表失败")?;

    let mut rows = range.rows();
    let header_row = rows.next().ok_or_else(|| anyhow!("Excel 内容为空"))?;

    let col_idx = header_row
        .iter()
        .position(|cell| cell_to_string(cell) == order_column_name)
        .ok_or_else(|| anyhow!("未找到订单列: {}", order_column_name))?;

    let mut order_rows = Vec::new();

    for (idx, row) in rows.enumerate() {
        let order = row
            .get(col_idx)
            .map(cell_to_string)
            .unwrap_or_default()
            .trim()
            .to_string();

        if order.is_empty() {
            continue;
        }

        order_rows.push(RowOrder {
            row_number: idx + 2,
            order_id: order,
        });
    }

    Ok(order_rows)
}

fn load_image_profiles(image_root_dir: &str) -> Result<Vec<ImageProfile>> {
    let root = Path::new(image_root_dir);
    if !root.exists() {
        return Err(anyhow!("图片根目录不存在: {}", image_root_dir));
    }
    if !root.is_dir() {
        return Err(anyhow!("图片根目录不是文件夹: {}", image_root_dir));
    }

    let mut profiles = Vec::new();

    for entry in fs::read_dir(root).context("读取图片根目录失败")? {
        let entry = entry.context("读取目录项失败")?;
        let path = entry.path();
        if !path.is_dir() {
            continue;
        }

        let folder_name = entry
            .file_name()
            .to_string_lossy()
            .trim()
            .to_string();
        if folder_name.is_empty() {
            continue;
        }

        let aliases: Vec<String> = folder_name
            .split("##")
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .collect();

        let mut images = Vec::new();
        for file in WalkDir::new(&path).into_iter().filter_map(|e| e.ok()) {
            if !file.file_type().is_file() {
                continue;
            }
            let p = file.path();
            let ext = p
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or_default()
                .to_lowercase();
            if ["png", "jpg", "jpeg", "webp", "bmp", "gif"].contains(&ext.as_str()) {
                images.push(p.to_path_buf());
            }
        }

        if images.is_empty() {
            continue;
        }

        profiles.push(ImageProfile {
            folder_name,
            aliases,
            images,
        });
    }

    if profiles.is_empty() {
        return Err(anyhow!("图片目录下没有可用图片文件夹"));
    }

    Ok(profiles)
}

fn match_profile<'a>(product_name: &str, profiles: &'a [ImageProfile]) -> Option<&'a ImageProfile> {
    let normalized_product = normalize_text(product_name);
    if normalized_product.is_empty() {
        return None;
    }

    let mut best: Option<(&ImageProfile, usize)> = None;

    for profile in profiles {
        for alias in &profile.aliases {
            let normalized_alias = normalize_text(alias);
            if normalized_alias.is_empty() {
                continue;
            }
            if normalized_product.contains(&normalized_alias) {
                let alias_len = normalized_alias.chars().count();
                match best {
                    Some((_, best_len)) if best_len >= alias_len => {}
                    _ => best = Some((profile, alias_len)),
                }
            }
        }
    }

    best.map(|(profile, _)| profile)
}

fn build_prompt(template: &str, product_name: &str) -> String {
    if template.contains("{product_name}") {
        template.replace("{product_name}", product_name)
    } else {
        format!("{}\n需要评价的商品是：{}。", template, product_name)
    }
}

fn call_ai(client: &Client, settings: &AppSettings, prompt: &str) -> Result<String> {
    if settings.ai_api_key.trim().is_empty() {
        return Err(anyhow!("AI API Key 为空"));
    }

    let mut base = settings.ai_api_base.trim().trim_end_matches('/').to_string();
    if !base.ends_with("/chat/completions") {
        base = format!("{}/chat/completions", base);
    }

    let mut headers = HeaderMap::new();
    headers.insert(USER_AGENT, HeaderValue::from_static("tauri-rating-assistant/1.0"));
    headers.insert(CONTENT_TYPE, HeaderValue::from_static("application/json"));

    let token = format!("Bearer {}", settings.ai_api_key.trim());
    headers.insert(
        AUTHORIZATION,
        HeaderValue::from_str(&token).context("AI API Key 包含非法字符")?,
    );

    let body = OpenAiChatRequest {
        model: settings.ai_model.trim().to_string(),
        messages: vec![OpenAiMessage {
            role: "user".to_string(),
            content: prompt.to_string(),
        }],
        temperature: 0.95,
    };

    let response = client
        .post(base)
        .headers(headers)
        .json(&body)
        .send()
        .context("请求 AI 接口失败")?;

    let text = ensure_success(response)?.text().context("读取 AI 响应失败")?;
    let json: serde_json::Value =
        serde_json::from_str(&text).with_context(|| format!("AI 响应非 JSON: {}", text))?;

    let content = json["choices"]
        .as_array()
        .and_then(|arr| arr.first())
        .and_then(|choice| choice["message"]["content"].as_str())
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .ok_or_else(|| anyhow!("AI 返回内容为空: {}", text))?;

    Ok(content)
}

fn write_summary_xlsx(path: &Path, rows: &[SummaryItem]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();

    let headers = ["订单号", "商品名称", "匹配图片目录", "评价文件", "图片数量", "状态"];
    for (col, header) in headers.iter().enumerate() {
        sheet
            .write_string(0, col as u16, *header)
            .context("写入表头失败")?;
    }

    for (idx, row) in rows.iter().enumerate() {
        let line = (idx + 1) as u32;
        sheet
            .write_string(line, 0, &row.order_id)
            .context("写入订单号失败")?;
        sheet
            .write_string(line, 1, &row.product_name)
            .context("写入商品名称失败")?;
        sheet
            .write_string(line, 2, &row.matched_folder)
            .context("写入目录失败")?;
        sheet
            .write_string(line, 3, &row.review_file)
            .context("写入评价文件失败")?;
        sheet
            .write_number(line, 4, row.image_count as f64)
            .context("写入图片数量失败")?;
        sheet
            .write_string(line, 5, &row.status)
            .context("写入状态失败")?;
    }

    workbook
        .save(path)
        .with_context(|| format!("保存汇总表失败: {}", path.display()))
}

fn unique_order_list(order_rows: &[RowOrder]) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut orders = Vec::new();

    for row in order_rows {
        if seen.insert(row.order_id.clone()) {
            orders.push(row.order_id.clone());
        }
    }

    orders
}

fn choose_random_images(images: &[PathBuf], count: usize) -> Vec<PathBuf> {
    if images.is_empty() || count == 0 {
        return Vec::new();
    }

    let mut rng = rand::rng();
    let mut selected = Vec::with_capacity(count);

    for _ in 0..count {
        if let Some(chosen) = images.choose(&mut rng) {
            selected.push(chosen.clone());
        }
    }

    selected
}

fn run_rating_internal(request: RunRequest) -> Result<RunResult> {
    let settings = request.settings;

    if settings.jst_cookie.trim().is_empty() {
        return Err(anyhow!("聚水潭 Cookie 为空，请先填写并验证"));
    }

    if settings.image_root_dir.trim().is_empty() {
        return Err(anyhow!("图片根目录不能为空"));
    }

    let order_rows = collect_order_rows(&request.excel_path, settings.order_column_name.trim())?;
    if order_rows.is_empty() {
        return Err(anyhow!("Excel 中没有可用订单号"));
    }

    let unique_orders = unique_order_list(&order_rows);
    let client = build_http_client()?;
    let order_products = fetch_products_by_orders(&client, settings.jst_cookie.trim(), &unique_orders)?;

    let profiles = load_image_profiles(settings.image_root_dir.trim())?;

    let excel_path = Path::new(&request.excel_path);
    let excel_stem = excel_path
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or("rating_task");

    let output_root = if settings.output_root_dir.trim().is_empty() {
        excel_path
            .parent()
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| PathBuf::from("."))
    } else {
        PathBuf::from(settings.output_root_dir.trim())
    };

    fs::create_dir_all(&output_root).context("创建输出根目录失败")?;

    let task_name = format!(
        "{}_{}",
        sanitize_filename(excel_stem),
        Local::now().format("%Y%m%d_%H%M%S")
    );
    let output_dir = output_root.join(task_name);
    fs::create_dir_all(&output_dir).context("创建任务输出目录失败")?;

    let mut summary_rows = Vec::new();
    let mut missing_products = HashSet::new();
    let mut failed_items = Vec::new();
    let mut generated_reviews = 0;
    let mut total_products = 0;

    let mut generated_reviews_map: HashMap<String, Vec<String>> = HashMap::new();
    let mut generated_review_keys: HashMap<String, HashSet<String>> = HashMap::new();

    for row in &order_rows {
        let products = order_products
            .get(&row.order_id)
            .cloned()
            .unwrap_or_default();

        if products.is_empty() {
            summary_rows.push(SummaryItem {
                order_id: row.order_id.clone(),
                product_name: String::new(),
                matched_folder: String::new(),
                review_file: String::new(),
                image_count: 0,
                status: format!("订单第{}行：未查询到商品", row.row_number),
            });
            continue;
        }

        let order_dir = output_dir.join(sanitize_filename(&row.order_id));
        fs::create_dir_all(&order_dir).context("创建订单目录失败")?;

        for (product_idx, product_name) in products.iter().enumerate() {
            total_products += 1;
            let profile = match_profile(product_name, &profiles);
            let Some(profile) = profile else {
                missing_products.insert(product_name.clone());
                summary_rows.push(SummaryItem {
                    order_id: row.order_id.clone(),
                    product_name: product_name.clone(),
                    matched_folder: String::new(),
                    review_file: String::new(),
                    image_count: 0,
                    status: "未匹配到图片目录".to_string(),
                });
                continue;
            };

            let prompt = build_prompt(settings.review_prompt_template.trim(), product_name);
            let key = profile.folder_name.clone();
            let existing_reviews = generated_reviews_map.entry(key.clone()).or_default();
            let existing_keys = generated_review_keys.entry(key.clone()).or_default();

            let mut final_review = None;
            for _ in 0..5 {
                let full_prompt = if existing_reviews.is_empty() {
                    prompt.clone()
                } else {
                    let recent = existing_reviews
                        .iter()
                        .rev()
                        .take(3)
                        .cloned()
                        .collect::<Vec<String>>()
                        .into_iter()
                        .rev()
                        .collect::<Vec<String>>()
                        .join("\n");
                    format!(
                        "{}\n另外要求：本次评价文案不能与以下示例重复或仅改写几个字，请生成全新表达：\n{}",
                        prompt, recent
                    )
                };

                match call_ai(&client, &settings, &full_prompt) {
                    Ok(review) => {
                        let normalized = normalize_text(&review);
                        if normalized.is_empty() || existing_keys.contains(&normalized) {
                            continue;
                        }
                        final_review = Some(review);
                        break;
                    }
                    Err(_) => {
                        continue;
                    }
                }
            }

            let review_text = match final_review {
                Some(text) => text,
                None => {
                    let msg = format!("{} -> {}: 评价生成失败", row.order_id, product_name);
                    failed_items.push(msg.clone());
                    summary_rows.push(SummaryItem {
                        order_id: row.order_id.clone(),
                        product_name: product_name.clone(),
                        matched_folder: profile.folder_name.clone(),
                        review_file: String::new(),
                        image_count: 0,
                        status: msg,
                    });
                    continue;
                }
            };

            generated_reviews += 1;
            existing_reviews.push(review_text.clone());
            existing_keys.insert(normalize_text(&review_text));

            let safe_folder_name = sanitize_filename(&profile.folder_name);
            let review_file_name = format!("{}_{}.txt", product_idx + 1, safe_folder_name);
            let review_file_path = order_dir.join(&review_file_name);
            fs::write(&review_file_path, review_text).context("写入评价文件失败")?;

            let chosen_images = choose_random_images(
                &profile.images,
                std::cmp::max(1, settings.images_per_product),
            );

            let mut copied_count = 0;
            for (image_idx, image_path) in chosen_images.iter().enumerate() {
                let ext = image_path
                    .extension()
                    .and_then(|v| v.to_str())
                    .unwrap_or("png");
                let image_name = format!(
                    "{}_{}_{}.{}",
                    product_idx + 1,
                    safe_folder_name,
                    image_idx + 1,
                    ext
                );
                let target_path = order_dir.join(image_name);
                fs::copy(image_path, target_path).context("复制图片失败")?;
                copied_count += 1;
            }

            summary_rows.push(SummaryItem {
                order_id: row.order_id.clone(),
                product_name: product_name.clone(),
                matched_folder: profile.folder_name.clone(),
                review_file: review_file_name,
                image_count: copied_count,
                status: "成功".to_string(),
            });
        }
    }

    let summary_file = output_dir.join("summary.xlsx");
    write_summary_xlsx(&summary_file, &summary_rows)?;

    Ok(RunResult {
        output_dir: output_dir.to_string_lossy().to_string(),
        summary_file: summary_file.to_string_lossy().to_string(),
        total_rows: order_rows.len(),
        total_orders: unique_orders.len(),
        total_products,
        generated_reviews,
        missing_products: {
            let mut list = missing_products.into_iter().collect::<Vec<String>>();
            list.sort();
            list
        },
        failed_items,
    })
}

#[tauri::command]
fn load_settings(app: AppHandle) -> Result<AppSettings, String> {
    let path = settings_path(&app).map_err(|e| e.to_string())?;
    if !path.exists() {
        let settings = AppSettings::default();
        let content = serde_json::to_string_pretty(&settings).map_err(|e| e.to_string())?;
        fs::write(&path, content).map_err(|e| format!("初始化配置文件失败: {}", e))?;
        return Ok(settings);
    }

    let text = fs::read_to_string(&path).map_err(|e| format!("读取配置文件失败: {}", e))?;
    let mut settings: AppSettings =
        serde_json::from_str(&text).map_err(|e| format!("配置文件格式错误: {}", e))?;

    if settings.order_column_name.trim().is_empty() {
        settings.order_column_name = "订单号".to_string();
    }
    if settings.images_per_product == 0 {
        settings.images_per_product = 5;
    }

    Ok(settings)
}

#[tauri::command]
fn save_settings(app: AppHandle, settings: AppSettings) -> Result<(), String> {
    let path = settings_path(&app).map_err(|e| e.to_string())?;
    let content = serde_json::to_string_pretty(&settings).map_err(|e| e.to_string())?;
    fs::write(path, content).map_err(|e| format!("写入配置文件失败: {}", e))
}

#[tauri::command]
fn open_jushuitan_login_window(app: AppHandle, login_url: Option<String>) -> Result<(), String> {
    let url = login_url
        .unwrap_or_else(|| "https://www.erp321.com/app/order/order/list.aspx?_c=jst-epaas".to_string())
        .trim()
        .to_string();
    if url.is_empty() {
        return Err("登录 URL 不能为空".to_string());
    }

    let parsed = Url::parse(&url).map_err(|e| format!("登录 URL 格式错误: {}", e))?;
    let label = "jst-login";

    if let Some(window) = app.get_webview_window(label) {
        window
            .navigate(parsed)
            .map_err(|e| format!("跳转登录页失败: {}", e))?;
        window.show().map_err(|e| format!("显示登录窗口失败: {}", e))?;
        window
            .set_focus()
            .map_err(|e| format!("聚焦登录窗口失败: {}", e))?;
        return Ok(());
    }

    #[cfg(target_os = "windows")]
    let builder = WebviewWindowBuilder::new(&app, label, WebviewUrl::External(parsed))
        .title("聚水潭登录")
        .inner_size(1280.0, 860.0)
        .resizable(true)
        // 某些站点在 Windows WebView2 上会因为 UA 识别导致白屏，显式伪装为桌面 Chrome。
        .user_agent(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36",
        );

    #[cfg(not(target_os = "windows"))]
    let builder = WebviewWindowBuilder::new(&app, label, WebviewUrl::External(parsed))
        .title("聚水潭登录")
        .inner_size(1280.0, 860.0)
        .resizable(true);

    builder
        .build()
        .map_err(|e| format!("创建登录窗口失败: {}", e))?;

    Ok(())
}

#[tauri::command]
fn capture_jushuitan_cookie(app: AppHandle) -> Result<String, String> {
    let window = app
        .get_webview_window("jst-login")
        .ok_or_else(|| "未找到登录窗口，请先点击“打开登录窗口”并完成登录".to_string())?;

    let mut all_cookies = window
        .cookies()
        .map_err(|e| format!("读取登录窗口 Cookie 失败: {}", e))?;

    // 仅保留聚水潭相关域名 Cookie，避免把无关站点 Cookie 混入请求头。
    all_cookies.retain(|cookie| {
        cookie
            .domain()
            .map(|domain| domain.contains("erp321.com"))
            .unwrap_or(false)
    });

    if all_cookies.is_empty() {
        let urls = [
            "https://www.erp321.com/",
            "https://src.erp321.com/",
            "https://apiweb.erp321.com/",
            "https://ss.erp321.com/",
            "https://api.erp321.com/",
        ];
        for raw_url in urls {
            if let Ok(parsed) = Url::parse(raw_url) {
                if let Ok(mut cookies) = window.cookies_for_url(parsed) {
                    all_cookies.append(&mut cookies);
                }
            }
        }
        all_cookies.retain(|cookie| {
            cookie
                .domain()
                .map(|domain| domain.contains("erp321.com"))
                .unwrap_or(false)
        });
    }

    if all_cookies.is_empty() {
        return Err("未提取到聚水潭 Cookie，请确认已在登录窗口完成登录".to_string());
    }

    let mut cookie_map = HashMap::<String, String>::new();
    for cookie in all_cookies {
        cookie_map.insert(cookie.name().to_string(), cookie.value().to_string());
    }
    let mut pairs = cookie_map.into_iter().collect::<Vec<(String, String)>>();
    pairs.sort_by(|a, b| a.0.cmp(&b.0));
    let cookie_header = pairs
        .into_iter()
        .map(|(k, v)| format!("{k}={v}"))
        .collect::<Vec<String>>()
        .join("; ");

    if cookie_header.trim().is_empty() {
        return Err("提取到的 Cookie 为空".to_string());
    }

    let client = build_http_client().map_err(|e| e.to_string())?;
    get_view_state(&client, &cookie_header).map_err(|e| format!("Cookie 提取成功，但校验失败: {}", e))?;

    Ok(cookie_header)
}

#[tauri::command]
fn validate_jushuitan_cookie(request: ValidateCookieRequest) -> Result<ValidateCookieResult, String> {
    if request.cookie.trim().is_empty() {
        return Ok(ValidateCookieResult {
            valid: false,
            message: "Cookie 为空".to_string(),
        });
    }

    let client = build_http_client().map_err(|e| e.to_string())?;

    match get_view_state(&client, request.cookie.trim()) {
        Ok(_) => Ok(ValidateCookieResult {
            valid: true,
            message: "Cookie 可用".to_string(),
        }),
        Err(err) => Ok(ValidateCookieResult {
            valid: false,
            message: err.to_string(),
        }),
    }
}

#[tauri::command]
fn test_review_prompt(request: TestPromptRequest) -> Result<TestPromptResult, String> {
    let settings = request.settings;
    if settings.ai_api_key.trim().is_empty() {
        return Err("AI API Key 为空，请先在第二步配置".to_string());
    }
    if settings.ai_model.trim().is_empty() {
        return Err("AI 模型为空，请先在第二步配置".to_string());
    }
    if settings.review_prompt_template.trim().is_empty() {
        return Err("提示词模板为空，请先在第二步配置".to_string());
    }

    let product_name = request
        .product_name
        .unwrap_or_else(|| "示例商品".to_string())
        .trim()
        .to_string();
    let product_name = if product_name.is_empty() {
        "示例商品".to_string()
    } else {
        product_name
    };

    let prompt = build_prompt(settings.review_prompt_template.trim(), &product_name);
    let client = build_http_client().map_err(|e| e.to_string())?;
    let review = call_ai(&client, &settings, &prompt).map_err(|e| e.to_string())?;

    Ok(TestPromptResult { prompt, review })
}

#[tauri::command]
async fn run_rating_task(request: RunRequest) -> Result<RunResult, String> {
    tauri::async_runtime::spawn_blocking(move || run_rating_internal(request).map_err(|e| e.to_string()))
        .await
        .map_err(|e| format!("任务执行线程异常: {}", e))?
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            load_settings,
            save_settings,
            open_jushuitan_login_window,
            capture_jushuitan_cookie,
            validate_jushuitan_cookie,
            test_review_prompt,
            run_rating_task
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
