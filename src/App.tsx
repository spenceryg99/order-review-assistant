import { useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import {
  Alert,
  Button,
  Card,
  ConfigProvider,
  Form,
  Input,
  InputNumber,
  Progress,
  Space,
  Steps,
  Tag,
  Typography,
} from "antd";
import {
  FileImageOutlined,
  RobotOutlined,
  SafetyCertificateOutlined,
  FileExcelOutlined,
} from "@ant-design/icons";
import "antd/dist/reset.css";
import "./App.css";

const { Title, Paragraph, Text } = Typography;

type AppSettings = {
  apiBaseUrl: string;
  jstLoginUrl: string;
  jstCookie: string;
  jstOwnerCoId: string;
  jstAuthorizeCoId: string;
  jstUid: string;
  aiApiBase: string;
  aiApiKey: string;
  aiModel: string;
  reviewPromptTemplate: string;
  imageRootDir: string;
  outputRootDir: string;
  orderColumnName: string;
  imagesPerProduct: number;
};

type ValidateCookieResult = {
  valid: boolean;
  message: string;
};

type PromptTestResult = {
  prompt: string;
  review: string;
};

type RunResult = {
  outputDir: string;
  summaryFile: string;
  totalRows: number;
  totalOrders: number;
  totalProducts: number;
  generatedReviews: number;
  missingProducts: string[];
  failedItems: string[];
};

type CookieState = "unknown" | "valid" | "invalid";

const defaultSettings: AppSettings = {
  apiBaseUrl: "http://192.168.1.166/index.php/api/Yingdao/",
  jstLoginUrl: "https://www.erp321.com/",
  jstCookie: "",
  jstOwnerCoId: "14805587",
  jstAuthorizeCoId: "14805587",
  jstUid: "21419081",
  aiApiBase: "https://dashscope.aliyuncs.com/compatible-mode/v1",
  aiApiKey: "",
  aiModel: "qwen-plus",
  reviewPromptTemplate:
    "请生成1条符合真人评价特点的淘宝商品好评，需要评价的商品是：{product_name}。直接返回评价内容，不要其他任何描述文字。",
  imageRootDir: "",
  outputRootDir: "",
  orderColumnName: "订单号",
  imagesPerProduct: 5,
};

const loadingPhases = ["读取 Excel", "查询订单商品", "匹配图片与生成评价", "写入结果与汇总"];

function getStep1Missing(settings: AppSettings): string[] {
  const missing: string[] = [];
  if (!settings.imageRootDir.trim()) missing.push("图片目录");
  return missing;
}

function getStep2Missing(settings: AppSettings): string[] {
  const missing: string[] = [];
  if (!settings.aiApiBase.trim()) missing.push("AI API Base");
  if (!settings.aiApiKey.trim()) missing.push("AI API Key");
  if (!settings.aiModel.trim()) missing.push("AI 模型");
  if (!settings.reviewPromptTemplate.trim()) missing.push("提示词模板");
  return missing;
}

function getSuggestedStep(step1Ready: boolean, step2Ready: boolean, step3Ready: boolean): number {
  if (!step1Ready) return 0;
  if (!step2Ready) return 1;
  if (!step3Ready) return 2;
  return 3;
}

function formatSeconds(value: number): string {
  const m = Math.floor(value / 60);
  const s = value % 60;
  if (m === 0) return `${s}s`;
  return `${m}m ${s.toString().padStart(2, "0")}s`;
}

function App() {
  const [settings, setSettings] = useState<AppSettings>(defaultSettings);
  const [excelPath, setExcelPath] = useState("");
  const [currentStep, setCurrentStep] = useState(0);

  const [running, setRunning] = useState(false);
  const [runElapsed, setRunElapsed] = useState(0);
  const [runProgress, setRunProgress] = useState(0);
  const [bootChecking, setBootChecking] = useState(true);

  const [savingStep1, setSavingStep1] = useState(false);
  const [savingStep2, setSavingStep2] = useState(false);
  const [promptTesting, setPromptTesting] = useState(false);
  const [cookieChecking, setCookieChecking] = useState(false);
  const [cookieExtracting, setCookieExtracting] = useState(false);
  const [cookieClosingWindow, setCookieClosingWindow] = useState(false);
  const [cookieResettingProfile, setCookieResettingProfile] = useState(false);
  const [cookieDiagLoading, setCookieDiagLoading] = useState(false);
  const [cookieDiagnostics, setCookieDiagnostics] = useState("");

  const [cookieState, setCookieState] = useState<CookieState>("unknown");
  const [cookieStatus, setCookieStatus] = useState("");

  const [promptSampleProduct, setPromptSampleProduct] = useState("老面包");
  const [promptTestResult, setPromptTestResult] = useState<PromptTestResult | null>(null);

  const [result, setResult] = useState<RunResult | null>(null);
  const [error, setError] = useState("");

  const step1Missing = useMemo(() => getStep1Missing(settings), [settings]);
  const step2Missing = useMemo(() => getStep2Missing(settings), [settings]);

  const step1Ready = step1Missing.length === 0;
  const step2Ready = step2Missing.length === 0;
  const step3Ready = cookieState === "valid";
  const step4Ready = step1Ready && step2Ready && step3Ready;

  const phaseText = useMemo(() => loadingPhases[Math.floor(runElapsed / 6) % loadingPhases.length], [runElapsed]);

  useEffect(() => {
    (async () => {
      setBootChecking(true);
      try {
        const loaded = await invoke<AppSettings>("load_settings");
        setSettings(loaded);

        const s1Ready = getStep1Missing(loaded).length === 0;
        const s2Ready = getStep2Missing(loaded).length === 0;

        let cState: CookieState = "unknown";
        let cMsg = "请先完成前两步";

        if (!s1Ready || !s2Ready) {
          cState = "unknown";
          cMsg = "等待前两步配置完成后再检测 Cookie";
        } else if (!loaded.jstCookie.trim()) {
          cState = "invalid";
          cMsg = "未发现 Cookie，请执行第三步";
        } else {
          const res = await invoke<ValidateCookieResult>("validate_jushuitan_cookie", {
            request: { cookie: loaded.jstCookie },
          });
          if (res.valid) {
            cState = "valid";
            cMsg = "Cookie 可用";
          } else {
            cState = "invalid";
            cMsg = `Cookie 失效：${res.message}`;
          }
        }

        setCookieState(cState);
        setCookieStatus(cMsg);
        setCurrentStep(getSuggestedStep(s1Ready, s2Ready, cState === "valid"));
      } catch (err) {
        setError(String(err));
      } finally {
        setBootChecking(false);
      }
    })();
  }, []);

  useEffect(() => {
    if (!running) return;
    const timer = window.setInterval(() => setRunElapsed((prev) => prev + 1), 1000);
    return () => window.clearInterval(timer);
  }, [running]);

  useEffect(() => {
    if (!running) return;
    const timer = window.setInterval(() => {
      setRunProgress((prev) => {
        if (prev >= 92) return prev;
        const delta = Math.max(1, Math.floor((100 - prev) / 10));
        return Math.min(92, prev + delta);
      });
    }, 800);
    return () => window.clearInterval(timer);
  }, [running]);

  const updateSettings = <K extends keyof AppSettings>(key: K, value: AppSettings[K]) => {
    setSettings((prev) => ({ ...prev, [key]: value }));
  };

  const saveSettings = async (next: AppSettings = settings) => {
    await invoke("save_settings", { settings: next });
  };

  const validateCookieByValue = async (cookie: string, withLoading = true): Promise<boolean> => {
    if (withLoading) setCookieChecking(true);
    try {
      if (!cookie.trim()) {
        setCookieState("invalid");
        setCookieStatus("未检测到 Cookie，请先登录并提取");
        return false;
      }

      const res = await invoke<ValidateCookieResult>("validate_jushuitan_cookie", {
        request: { cookie },
      });

      if (res.valid) {
        setCookieState("valid");
        setCookieStatus(`校验通过：${res.message}`);
        return true;
      }

      setCookieState("invalid");
      setCookieStatus(`校验失败：${res.message}`);
      return false;
    } catch (err) {
      setCookieState("invalid");
      setCookieStatus(`校验异常：${String(err)}`);
      return false;
    } finally {
      if (withLoading) setCookieChecking(false);
    }
  };

  const pickImageRoot = async () => {
    const selected = await open({ multiple: false, directory: true });
    if (typeof selected === "string") updateSettings("imageRootDir", selected);
  };

  const pickOutputRoot = async () => {
    const selected = await open({ multiple: false, directory: true });
    if (typeof selected === "string") updateSettings("outputRootDir", selected);
  };

  const pickExcel = async () => {
    const selected = await open({
      multiple: false,
      directory: false,
      filters: [{ name: "Excel", extensions: ["xlsx", "xls", "xlsm", "xlsb"] }],
    });
    if (typeof selected === "string") setExcelPath(selected);
  };

  const saveStep1 = async () => {
    setSavingStep1(true);
    setError("");
    try {
      if (!step1Ready) throw new Error(`请补全：${step1Missing.join("、")}`);
      await saveSettings();
    } catch (err) {
      setError(String(err));
    } finally {
      setSavingStep1(false);
    }
  };

  const saveStep2 = async () => {
    setSavingStep2(true);
    setError("");
    try {
      if (!step2Ready) throw new Error(`请补全：${step2Missing.join("、")}`);
      await saveSettings();

      if (settings.jstCookie.trim()) {
        await validateCookieByValue(settings.jstCookie, false);
      }
    } catch (err) {
      setError(String(err));
    } finally {
      setSavingStep2(false);
    }
  };

  const testPrompt = async () => {
    setPromptTesting(true);
    setError("");
    setPromptTestResult(null);

    try {
      if (!step2Ready) throw new Error(`请先补全第二步：${step2Missing.join("、")}`);

      const res = await invoke<PromptTestResult>("test_review_prompt", {
        request: {
          settings,
          productName: promptSampleProduct,
        },
      });
      setPromptTestResult(res);
    } catch (err) {
      setError(String(err));
    } finally {
      setPromptTesting(false);
    }
  };

  const openLoginWindow = async () => {
    setError("");
    try {
      await invoke("open_jushuitan_login_window", {
        loginUrl: settings.jstLoginUrl,
      });
    } catch (err) {
      setError(String(err));
    }
  };

  const extractCookie = async () => {
    setCookieExtracting(true);
    setCookieStatus("");
    setError("");

    try {
      const cookie = await invoke<string>("capture_jushuitan_cookie");
      const nextSettings = { ...settings, jstCookie: cookie };
      setSettings(nextSettings);
      await saveSettings(nextSettings);
      await validateCookieByValue(cookie, false);
    } catch (err) {
      setCookieState("invalid");
      setCookieStatus(`自动提取失败：${String(err)}`);
    } finally {
      setCookieExtracting(false);
    }
  };

  const closeLoginWindow = async () => {
    setCookieClosingWindow(true);
    setError("");
    try {
      await invoke("close_jushuitan_login_window");
      setCookieStatus("登录浏览器已关闭");
    } catch (err) {
      setError(String(err));
    } finally {
      setCookieClosingWindow(false);
    }
  };

  const resetLoginProfile = async () => {
    setCookieResettingProfile(true);
    setError("");
    try {
      const message = await invoke<string>("reset_jushuitan_login_webview_profile");
      setCookieStatus(message || "登录浏览器缓存已重置");
      setCookieDiagnostics("（已重置登录浏览器缓存，可重新打开登录浏览器）");
    } catch (err) {
      setError(String(err));
    } finally {
      setCookieResettingProfile(false);
    }
  };

  const loadCookieDiagnostics = async () => {
    setCookieDiagLoading(true);
    setError("");
    try {
      const timeoutMs = 8000;
      const text = await Promise.race<string>([
        invoke<string>("get_jushuitan_login_diagnostics"),
        new Promise<string>((_, reject) =>
          window.setTimeout(() => reject(new Error(`读取诊断日志超时（>${timeoutMs / 1000}s）`)), timeoutMs),
        ),
      ]);
      setCookieDiagnostics(text);
    } catch (err) {
      setError(String(err));
    } finally {
      setCookieDiagLoading(false);
    }
  };

  const clearCookieDiagnostics = async () => {
    setCookieDiagLoading(true);
    setError("");
    try {
      await invoke("clear_jushuitan_login_diagnostics");
      setCookieDiagnostics("（已清空）");
    } catch (err) {
      setError(String(err));
    } finally {
      setCookieDiagLoading(false);
    }
  };

  const runTask = async () => {
    setRunning(true);
    setRunElapsed(0);
    setRunProgress(8);
    setResult(null);
    setError("");

    try {
      if (!step4Ready) throw new Error("请先完成前 3 步");
      if (!excelPath.trim()) throw new Error("请先选择 Excel 文件");

      // 先让 React 完成一次渲染，确保“生成中/进度条”先显示出来。
      await new Promise<void>((resolve) => {
        requestAnimationFrame(() => resolve());
      });

      await saveSettings();

      const runResult = await invoke<RunResult>("run_rating_task", {
        request: {
          settings,
          excelPath,
        },
      });
      setResult(runResult);
      setRunProgress(100);
    } catch (err) {
      setError(String(err));
    } finally {
      await new Promise((resolve) => setTimeout(resolve, 350));
      setRunning(false);
      setRunProgress(0);
    }
  };

  const stepItems = [
    {
      title: "步骤 1",
      subTitle: step1Ready ? "已完成" : "待配置",
      description: "图片目录",
      icon: <FileImageOutlined />,
    },
    {
      title: "步骤 2",
      subTitle: step2Ready ? "已完成" : "待配置",
      description: "AI 与提示词",
      icon: <RobotOutlined />,
    },
    {
      title: "步骤 3",
      subTitle: step3Ready ? "可用" : "待检测",
      description: "Cookie",
      icon: <SafetyCertificateOutlined />,
    },
    {
      title: "步骤 4",
      subTitle: step4Ready ? "可执行" : "未解锁",
      description: "Excel 生成",
      icon: <FileExcelOutlined />,
    },
  ];

  return (
    <ConfigProvider
      theme={{
        token: {
          colorPrimary: "#176b8f",
          borderRadius: 10,
        },
      }}
    >
      <main className="app-page">
        <div className="app-shell">
          <Card bordered={false} className="top-card">
            <Space direction="vertical" size={2}>
              <Title level={3} style={{ margin: 0 }}>
                订单评价助手
              </Title>
              <Paragraph type="secondary" style={{ margin: 0 }}>
                使用 4 步流程完成配置并执行任务，顶部导航可点击定位。
              </Paragraph>
            </Space>
            <Steps
              className="step-nav"
              type="navigation"
              size="small"
              current={currentStep}
              onChange={(step) => setCurrentStep(step)}
              items={stepItems}
            />
          </Card>

          {bootChecking && (
            <Alert
              style={{ marginTop: 12 }}
              type="info"
              message="正在检测历史配置与 Cookie 状态..."
              showIcon
            />
          )}

          {error && (
            <Alert
              style={{ marginTop: 12 }}
              type="error"
              message={error}
              showIcon
              closable
              onClose={() => setError("")}
            />
          )}

          <div className="step-content">
            {currentStep === 0 && (
              <Card title="步骤 1：配置图片目录" className="content-card">
                <Form layout="vertical">
                  <Form.Item label="图片目录" required>
                    <Space.Compact style={{ width: "100%" }}>
                      <Input
                        value={settings.imageRootDir}
                        onChange={(e) => updateSettings("imageRootDir", e.target.value)}
                        placeholder="选择本地图片分类目录"
                      />
                      <Button onClick={pickImageRoot}>选择目录</Button>
                    </Space.Compact>
                  </Form.Item>

                  {!step1Ready && (
                    <Alert type="warning" showIcon message={`缺少：${step1Missing.join("、")}`} style={{ marginBottom: 12 }} />
                  )}

                  <Button type="primary" loading={savingStep1} onClick={saveStep1}>
                    保存第一步
                  </Button>
                </Form>
              </Card>
            )}

            {currentStep === 1 && (
              <Card title="步骤 2：配置 AI 与提示词" className="content-card">
                {!step1Ready && (
                  <Alert type="warning" showIcon message="请先完成第一步（图片目录）" style={{ marginBottom: 12 }} />
                )}

                <Form layout="vertical">
                  <Form.Item label="AI API Base" required>
                    <Input
                      value={settings.aiApiBase}
                      onChange={(e) => updateSettings("aiApiBase", e.target.value)}
                      disabled={!step1Ready}
                    />
                  </Form.Item>

                  <div className="grid-two">
                    <Form.Item label="AI 模型" required>
                      <Input
                        value={settings.aiModel}
                        onChange={(e) => updateSettings("aiModel", e.target.value)}
                        disabled={!step1Ready}
                      />
                    </Form.Item>

                    <Form.Item label="AI API Key" required>
                      <Input.Password
                        value={settings.aiApiKey}
                        onChange={(e) => updateSettings("aiApiKey", e.target.value)}
                        disabled={!step1Ready}
                      />
                    </Form.Item>
                  </div>

                  <Form.Item label="提示词模板（支持 {product_name}）" required>
                    <Input.TextArea
                      rows={5}
                      value={settings.reviewPromptTemplate}
                      onChange={(e) => updateSettings("reviewPromptTemplate", e.target.value)}
                      disabled={!step1Ready}
                    />
                  </Form.Item>

                  <Card size="small" title="测试提示词" className="inner-card">
                    <Space direction="vertical" style={{ width: "100%" }}>
                      <Space.Compact style={{ width: "100%" }}>
                        <Input
                          value={promptSampleProduct}
                          onChange={(e) => setPromptSampleProduct(e.target.value)}
                          disabled={!step1Ready}
                          placeholder="输入测试商品名"
                        />
                        <Button loading={promptTesting} onClick={testPrompt} disabled={!step1Ready || running}>
                          测试
                        </Button>
                      </Space.Compact>

                      {promptTestResult && (
                        <Alert
                          type="success"
                          showIcon
                          message="测试返回"
                          description={<Text>{promptTestResult.review}</Text>}
                        />
                      )}
                    </Space>
                  </Card>

                  {!step2Ready && step1Ready && (
                    <Alert type="warning" showIcon message={`缺少：${step2Missing.join("、")}`} style={{ marginTop: 12 }} />
                  )}

                  <div style={{ marginTop: 12 }}>
                    <Button type="primary" loading={savingStep2} onClick={saveStep2} disabled={!step1Ready}>
                      保存第二步
                    </Button>
                  </div>
                </Form>
              </Card>
            )}

            {currentStep === 2 && (
              <Card title="步骤 3：配置并检测 Cookie" className="content-card">
                {(!step1Ready || !step2Ready) && (
                  <Alert type="warning" showIcon message="请先完成前两步配置" style={{ marginBottom: 12 }} />
                )}

                <Space direction="vertical" style={{ width: "100%" }}>
                  <Form layout="vertical">
                    <Form.Item label="登录地址">
                      <Space.Compact style={{ width: "100%" }}>
                        <Input
                          value={settings.jstLoginUrl}
                          onChange={(e) => updateSettings("jstLoginUrl", e.target.value)}
                          disabled={!step1Ready || !step2Ready}
                        />
                        <Button onClick={openLoginWindow} disabled={!step1Ready || !step2Ready || running}>
                          打开登录浏览器
                        </Button>
                      </Space.Compact>
                    </Form.Item>

                    <Space wrap>
                      <Button
                        type="primary"
                        loading={cookieExtracting}
                        onClick={extractCookie}
                        disabled={!step1Ready || !step2Ready || running}
                      >
                        自动提取并保存
                      </Button>
                      <Button
                        loading={cookieChecking}
                        onClick={() => validateCookieByValue(settings.jstCookie)}
                        disabled={!step1Ready || !step2Ready || cookieExtracting || cookieClosingWindow || running}
                      >
                        重新校验
                      </Button>
                      <Button
                        danger
                        loading={cookieClosingWindow}
                        onClick={closeLoginWindow}
                        disabled={cookieExtracting || cookieChecking || cookieResettingProfile || running}
                      >
                        关闭登录浏览器
                      </Button>
                      <Button
                        loading={cookieResettingProfile}
                        onClick={resetLoginProfile}
                        disabled={cookieExtracting || cookieChecking || cookieClosingWindow || running}
                      >
                        重置登录浏览器缓存
                      </Button>
                      <Button
                        loading={cookieDiagLoading}
                        onClick={loadCookieDiagnostics}
                        disabled={cookieResettingProfile || running}
                      >
                        查看诊断日志
                      </Button>
                      <Button
                        loading={cookieDiagLoading}
                        onClick={clearCookieDiagnostics}
                        disabled={cookieResettingProfile || running}
                      >
                        清空诊断日志
                      </Button>
                      {cookieState === "valid" && <Tag color="success">Cookie 可用</Tag>}
                      {cookieState === "invalid" && <Tag color="error">Cookie 无效</Tag>}
                      {cookieState === "unknown" && <Tag>未检测</Tag>}
                    </Space>

                    <Form.Item label="当前 Cookie（自动写入）" style={{ marginTop: 12 }}>
                      <Input.TextArea rows={4} value={settings.jstCookie} readOnly />
                    </Form.Item>

                    <Form.Item label="登录浏览器诊断日志">
                      <Input.TextArea
                        rows={9}
                        value={cookieDiagnostics}
                        readOnly
                        placeholder="点击“查看诊断日志”查看外部浏览器登录与抓取 Cookie 事件"
                      />
                    </Form.Item>
                  </Form>

                  <Alert
                    type={cookieState === "valid" ? "success" : cookieState === "invalid" ? "warning" : "info"}
                    showIcon
                    message={cookieStatus || "等待操作"}
                  />
                </Space>
              </Card>
            )}

            {currentStep === 3 && (
              <Card title="步骤 4：选择 Excel 并生成" className="content-card">
                {!step4Ready && (
                  <Alert type="warning" showIcon message="请先完成前 3 步" style={{ marginBottom: 12 }} />
                )}

                {running && (
                  <Card size="small" className="inner-card" style={{ marginBottom: 12 }}>
                    <Space direction="vertical" style={{ width: "100%" }} size={6}>
                      <Text strong>正在生成，请勿关闭窗口</Text>
                      <Text type="secondary">当前阶段：{phaseText}</Text>
                      <Progress percent={runProgress} status="active" />
                      <Text type="secondary">已运行：{formatSeconds(runElapsed)}</Text>
                    </Space>
                  </Card>
                )}

                <Form layout="vertical">
                  <Form.Item label="Excel 文件" required>
                    <Space.Compact style={{ width: "100%" }}>
                      <Input
                        value={excelPath}
                        onChange={(e) => setExcelPath(e.target.value)}
                        disabled={!step4Ready}
                        placeholder="选择本次要处理的 Excel"
                      />
                      <Button onClick={pickExcel} disabled={!step4Ready || running}>
                        选择文件
                      </Button>
                    </Space.Compact>
                  </Form.Item>

                  <div className="grid-two">
                    <Form.Item label="订单列名">
                      <Input
                        value={settings.orderColumnName}
                        onChange={(e) => updateSettings("orderColumnName", e.target.value)}
                        disabled={!step4Ready}
                      />
                    </Form.Item>

                    <Form.Item label="每商品图片数">
                      <InputNumber
                        min={1}
                        max={20}
                        style={{ width: "100%" }}
                        value={settings.imagesPerProduct}
                        onChange={(value) => updateSettings("imagesPerProduct", Number(value) || 5)}
                        disabled={!step4Ready}
                      />
                    </Form.Item>
                  </div>

                  <Form.Item label="输出目录">
                    <Space.Compact style={{ width: "100%" }}>
                      <Input
                        value={settings.outputRootDir}
                        onChange={(e) => updateSettings("outputRootDir", e.target.value)}
                        disabled={!step4Ready}
                        placeholder="可留空（默认 Excel 同目录）"
                      />
                      <Button onClick={pickOutputRoot} disabled={!step4Ready || running}>
                        选择目录
                      </Button>
                    </Space.Compact>
                  </Form.Item>

                  <Button type="primary" size="large" onClick={runTask} disabled={!step4Ready || running || !excelPath.trim()}>
                    {running ? "生成中..." : "立即生成"}
                  </Button>
                </Form>

                {result && (
                  <Card size="small" title="执行结果" className="inner-card" style={{ marginTop: 14 }}>
                    <Space direction="vertical" size={4}>
                      <Text>
                        共 {result.totalRows} 行，{result.totalOrders} 个订单，{result.totalProducts} 个商品，成功生成 {result.generatedReviews} 条评价。
                      </Text>
                      <Text copyable>输出目录：{result.outputDir}</Text>
                      <Text copyable>汇总文件：{result.summaryFile}</Text>
                    </Space>
                  </Card>
                )}
              </Card>
            )}
          </div>
        </div>
      </main>
    </ConfigProvider>
  );
}

export default App;
