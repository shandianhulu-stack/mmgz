# MIMO 工资评分生成器 Final V1.1 并入说明

## 1. 当前定位
独立 HTML 本地版，后续并入 MIMO 辅助软件。

## 2. 核心接口
- `GET /api/health`
- `POST /api/wage-score/preview`
- `POST /api/wage-score/generate`

## 3. 表单字段
- `score_file`：米墨评分 Excel，必传。
- `settlement_file`：结算过程 / 工资明细表，建议传。
- `strategy`：评分策略。
- `output_name`：导出文件名。

## 4. 必须保留的 strategy
- `human_like`：人工少量调分，严格不超发，默认。
- `human_batch_balance`：批次差值贴0，允许微超几分钱。
- `closest_allow_overpay`：逐人四舍五入，单人差值最小，可能微超。
- `zero_first_high`：多行均衡高分，旧 V0.6。
- `last_fill`：最后一场补差，旧逻辑。

## 5. 默认策略
默认必须是 `human_like`，因为用户已经选择 B：严格不超发。
其他策略保留为运营/财务对比使用，不得删除。

## 6. 算法主线
每场金额 700 元；评分精度为百分比小数点后 2 位；最小金额单位为 0.07 元。
`human_like` 的金额层严格向下贴近目标，不允许任何人微超；评分层按人工习惯：多个 100%，最后一行过低才拆成 2-5 行，优先使用 90%、85%、80%、75%、70% 等整数档。

## 7. 页面要求
- 保留 Sheet1 预览。
- 保留最终金额明细预览，默认全部展开。
- 保留策略下拉切换。
- 生成前必须自动刷新预览。
