# Investment Thesis Toolkit

투자 분석·의사결정 워크플로우를 자동화하는 Claude Skill 모음. Claude Code, Cowork 등 Claude 기반 환경에서 슬래시 명령어와 자연어 요청으로 활용한다.

## 수록 스킬

### [korean-stock-portfolio](korean-stock-portfolio/)
한국 주식 포트폴리오를 분기별로 분석·관리한다. 시장 데이터 수집 → 12M Forward EPS 기반 4단계 시나리오(보수/기준/낙관/AI 재평가) 적정가 산정 → 6개 시트 엑셀 산출 → 차기 실적 발표일 캘린더 등록까지 일관 수행.

**슬래시 명령어 5종**

| 명령어 | 용도 |
|--------|------|
| `/적정가산정 [종목]` | 4단계 시나리오로 종합 적정가 평가 |
| `/실적분석 [종목] [분기]` | 분기 실적 반영, 적정가 재산정, 히스토리 누적 |
| `/실적프리뷰 [종목]` | 차기 실적 발표 전 시나리오 분석 |
| `/논리점검 [종목]` | 분기별 히스토리로 투자 논리 추적 |
| `/ETF분석 [ETF명]` | 편입 종목 가중평균으로 ETF 적정가 산출 |

상세 사용법은 [korean-stock-portfolio/SKILL.md](korean-stock-portfolio/SKILL.md) 참조.

## 사용 방법

### 1. Claude Code / Cowork에서 스킬 등록
스킬 디렉토리(예: `korean-stock-portfolio/`)를 Claude의 skills 경로에 배치하면 자동 로드된다.

```bash
# 디렉토리 형태로 사용
git clone https://github.com/solo-starter/investment-thesis-toolkit.git
cp -r investment-thesis-toolkit/korean-stock-portfolio ~/.claude/skills/

# 또는 .skill 번들 파일로 사용
# (각 스킬 디렉토리를 zip으로 묶어 .skill 확장자로 배포 가능)
```

### 2. 트리거
- **자연어**: "포트폴리오 업데이트해줘", "삼성전자 적정가 분석"
- **슬래시 명령어**: `/적정가산정 SK하이닉스`, `/실적분석 효성중공업 2025Q4`

## 로드맵

이 레포는 단일 스킬에 한정되지 않는다. 투자·업무 자동화에 유용한 다양한 스킬을 점진적으로 추가할 예정이다.

- 추가 분석 도메인 (글로벌 주식, 채권, 매크로 지표 등)
- 리서치·보고 자동화 (위클리 브리핑, 섹터 리포트)
- 협업 워크플로우 (Cowork 환경 최적화 명령어)

## 라이선스

미정 (TBD)
