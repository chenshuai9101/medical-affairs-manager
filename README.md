# Medical Affairs Manager — 医学部经理AI工作台

产品无关的通用医学事务AI辅助系统。覆盖医学经理日常工作的核心场景：学术推广、科室会PPT、文献追踪、KOL管理、销售培训、反对意见应对。

## 适用人群

- 医学经理 / 医学联络官（MSL）
- 产品经理（已明确产品或跨产品线）
- 销售培训经理
- 市场准入与学术推广人员

## 核心能力

| 能力 | 说明 |
|------|------|
| 学术推广材料准备 | 科室会PPT、学术幻灯、DA文案 |
| 产品知识问答 | 结合内部资料+PubMed文献+CT.gov |
| 医药代表培训 | 推荐话术、反对意见模拟 |
| KOL管理 | 专家画像、维护计划、学术支持 |
| 竞争分析 | 竞品追踪、差异化定位 |
| 文献速查 | 自动搜索PubMed+ClinicalTrials |

## 目录结构

```
├── SKILL.md                          # 主技能描述 + AI工作流
├── references/                       # 泛行业知识库
│   ├── medical-affairs-101.md        # 医学事务基础框架
│   ├── academic-promotion.md         # 学术推广方法论
│   ├── objection-handling.md         # 反对意见应对框架
│   ├── dept-meeting-template.md      # 科室会标准流程
│   └── literature-tracking.md        # 文献追踪流程
├── scripts/
│   └── generate_dept_meeting_ppt.py  # 通用科室会PPT生成器
└── assets/                           # 通用素材模板
    └── README.md
```
