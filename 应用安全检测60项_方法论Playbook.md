# 应用安全检测（60 项）真实实测 + DOCX 报告生成

面向 Web/内网应用的「60 项应用安全检测」标准化流程：真实登录取证 + 黑盒实测 → 套用 V2 模板生成正式 DOCX/PDF 报告 → 按约定交付。适用于同一平台多环境（UAT/SIT/DEV/生产）或结构相似的新系统。

## 前置输入（开工前先确认齐全）
- 目标：完整 URL（含 scheme/host/port/path），如 `https://mbmp-uat.dfmc.com.cn` 或 `http://10.124.9.253:443/fp/`
- 凭据：管理员账号 + 尽量再要一个普通用户账号（用于越权/数据隔离 PER-002/003 验证；缺则相关项标「未测试」）
- 交付去向：GitHub 仓库（如 `change1238/32` main）和/或本地目录（如 `D:\safetest`）
- 基准模板：既有 V2 DOCX 报告 + 对应改写脚本（`make_*.py`、`dump.py`）所在目录（默认 `/home/ubuntu/safetest`）
- 内网目标：确认本地是否可达、是否需 cpolar 隧道穿透

缺任何一项先用 message_user 一次性问清，不要臆测或用错凭据兜底。

## 内网穿透（仅当目标在内网/本地时）
- **优先 TCP 隧道，别用 HTTP 隧道**：cpolar HTTP 隧道会改写 Host 头、注入 `X-Forwarded-*`，很多后端网关据此拒绝 → 统一 401/502。TCP 隧道原样透传字节，请求与浏览器一致。
  ```
  cpolar tcp <IP>:<port> -config=<cpolar.yml>
  # 得到 tcp://x.tcp.cpolar.cn:PORT -> <IP>:<port>，用 http://x.tcp.cpolar.cn:PORT 访问
  ```
- 若拿到 401 且用户确认本地能正常访问 → 几乎必是 HTTP 隧道代理头问题，直接切 TCP 隧道。
- cpolar 免费版同一时刻只有一条隧道：文件服务隧道与应用隧道不能并存，按需切换。
- 交付内网结果默认「只改本地、不推 GitHub」（除非用户明确要推）。

## 阶段一：黑盒探测（curl，不需登录）
逐项取真实事实，别照抄模板旧结论：
- 端口/TLS：`http` vs `https`、端口、证书、TLS 版本（PORT）
- 响应头：HSTS / X-Frame-Options / CSP / X-Content-Type-Options / Server 版本泄露（BASIC/TRAN/VULN）
- CORS：`OPTIONS` 带 `Origin: https://evil.example` 看是否反射任意源 + `Allow-Credentials:true`（SESS-002）
- 版本/信息泄露端点：`/actuator*`、`/swagger-ui`、`/v2|v3/api-docs`、`/druid`、`version.json`（注意区分真实响应 vs SPA fallback 200）
- 前端 JS bundle：抓 `index.*.js`，grep 硬编码密钥/RSA 公私钥/AES key/内网 IP/baseURL（FRON-002/KEY-001）
- 登录接口：验证码是否开启（`captchaEnabled`）、空体提交、错误提示是否泛化（AUTH-001/005/ERR-001）
- HTTP 方法：PUT/DELETE/TRACE 是否放行（VULN）

## 阶段二：真实登录 + 认证态深测（浏览器 computer 工具）
- 用管理员凭据登录；特殊字符密码（如 `V{Fc~39m`）易被 type 动作丢字符 → 校验密码框长度，必要时用 JS 直接 set value + 触发 input 事件。
- 从 localStorage/cookie 取 Token，配合 curl 做 API 级枚举（`getInfo` 角色权限、`getRouters` 真实菜单路由、`user/list`、`role/list`、`monitor/*`）。
- 逐一取证并截图：
  - 管理控制台（应用/用户/角色管理）
  - 越权：普通用户访问管理员 URL（无普通账号则标未测试）
  - 数据脱敏：用户手机号/邮箱、业务数据税号/身份证/银行账号/地址/VIN 是否明文（DATA-002/003）；横向滚表格露出敏感列取证
  - 审计日志 schema：账号/客户端 IP/模块/操作/方法/状态/时间 + 记录数（AUD-001/002）
  - 会话：Token 存储位置、HttpOnly/Secure、登出失效、超时、每次登录换签（SESS）
- 截图存 `fp_imgs/` 等目录；黑盒结论可用 PIL 把真实 curl 输出渲染成终端风格证据图。

## 阶段三：套模板生成报告（改写脚本，别手改每格）
以既有环境 DOCX 为结构基准，写 `make_<env>.py`：
1. 全局文本替换：系统名、环境标识、域名、IP、用户/租户、技术栈、日期、凭据脚注。**慎用裸 "DEV" 之类替换**（会误伤版本号），改用精确标识串。
2. 逐项覆盖 `测试结果 / 截图说明 / 修复建议`；**`风险等级` 是项固有值一般不改**，只改测试结果。缺凭据/无法验证的项据实标「未测试」，不适用标「不涉及」。
3. 图片：按标题（caption）段落定位 blip 引用替换真实截图；先 PIL 压缩（PNG→JPEG，宽≤1300px，quality~78），避免 DOCX 膨胀（20MB→~1-2MB）。
4. 更新风险汇总表、优先级表、概述统计段、变更记录，保证「符合/不符合/未测试/不涉及」计数与逐项一致。
5. 用 `dump.py`（caps/tables 模式）核对表格索引与项→表映射；跑完 grep 残留旧标识（EUSD/mbmp/旧域名/AES-GCM/租户 等）。
6. `libreoffice --headless --convert-to pdf` 出预览，`pdftoppm` 抽几页人工核对封面、DATA/KEY 证据页、脚注是否渲染正确。

## 阶段四：交付
- GitHub：`cp` 报告进本地 clone（如 `/home/ubuntu/repo32`）→ `git add 指定文件`（**绝不 `git add .`，绝不提交 token/凭据**）→ commit 写清系统/IP/关键结论 → `git push origin main`。
- 本地：经文件服务隧道上传到用户目录，或让用户本地 `git pull`。
- 交付后：吊销会话中明文出现过的 PAT/凭据（安全提醒），standby 后续需求。

## 60 项 13 类速查
PORT 端口 / BASIC 基线 / FRON 前端 / TRAN 传输 / AUTH 身份鉴别 / PER 权限 / SESS 会话 / KEY 密钥 / ERR 异常处理 / BUS 业务 / AUD 日志审计 / DATA 数据 / VULN 常见漏洞（SQLi/XSS/弱口令等）。风险分级：严重 / 高危 / 中危。
