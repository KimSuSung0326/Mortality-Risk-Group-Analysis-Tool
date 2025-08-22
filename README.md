# EDSD Count Analyzer 🩺

**사망위험군(Dead Sign) 분석 툴**

엑셀(.xlsx) 파일을 불러와 EDSD 지표를 계산하고, 결과(sign1, sign2, sign3)를 추가하여 새 파일로 저장하는 Windows Forms 기반 C# 애플리케이션입니다.

## 📋 주요 기능

### 엑셀 파일 로드
지정한 폴더에서 .xlsx 파일들을 자동으로 불러와 선택 가능

### EDSD Count 분석
호흡수, 심박 감지, SpO2, 거리(range), 신호 강도(rssi) 데이터를 기반으로 사망위험군 조건 판별
- **sign1**: 20행 이동 평균 호흡수
- **sign2**: 위험 신호 조건 충족 여부 (0/1)
- **sign3**: 누적된 위험 신호 카운트

### 시간대별 집계
- **오전** (12시 이전) 위험 신호 개수
- **오후** (12시 이후) 위험 신호 개수

### 엑셀 자동 저장
- 계산된 결과를 원본 엑셀에 열 추가 후 저장
- **저장 경로**: `EDSD_Data/{yyyy-MM-dd}/원본파일명.xlsx`

### 진행 상태 표시
Circular Progress Bar를 통해 파일 처리 진행률 시각화

### 로그 기록
실행 로그를 `log/{yyyy-MM-dd}.log` 파일로 자동 저장

## 🏥 지원 병원 조건 로직

### yn (영남, 경산)
- **3층 병실(311 ~ 320)**: `range = 1.66` 조건 적용
- **그 외**: `range = 1.77` 조건 적용

### h, jj, gj (효사랑, 전남제일, 구미제일)
- `range = 1.44` 조건 적용

### 공통 조건
- `breath > 0`
- `heart_detection == 1`
- 현재 호흡 ≤ 85% × 이전 호흡 평균
- `SpO2 < 95` 또는 `breath < 10`
- `radar_rssi > 200000`

## 📂 프로젝트 구조

```
count_dead_sign/
├── MainForm.cs               # UI 및 메인 로직
├── SaveExcelForm.cs          # 분석 요약 저장 UI
├── RoundedButton.cs          # 커스텀 버튼 클래스
├── CircularProgressBar.cs    # 원형 프로그레스바
├── log/                      # 실행 로그 저장 폴더
└── EDSD_Data/               # 분석된 엑셀 저장 폴더
```

## ⚙️ 요구사항

- **운영체제**: Windows
- **.NET Framework**: 4.7.2 이상
- **NuGet 패키지**: EPPlus (OfficeOpenXml)

### 설치 예시
```bash
Install-Package EPPlus -Version 5.8.0
```

## 🚀 사용 방법

1. **프로그램 실행** (`count_dead_sign.exe`)
2. **폴더 선택** 버튼 → 분석할 `.xlsx` 파일들이 있는 폴더 선택
3. **분석 시작** 버튼 클릭
4. 진행 상황이 Circular Progress Bar에 표시됨
5. **완료 후 결과**는 다음 위치에 저장됨:
   - `/EDSD_Data/{날짜}/원본파일명.xlsx`
6. **실행 로그 확인**: `/log/{yyyy-MM-dd}.log`

## 📊 엑셀 출력 결과

### 추가되는 열:
- **sign1**: 최근 20행 평균 호흡수
- **sign2**: 위험 신호 여부 (1 = 위험, 0 = 정상)
- **sign3**: 누적 위험 신호 개수

> ⚠️ **위험 신호**(sign2=1)일 경우 해당 셀은 **빨간색**으로 표시됩니다.
