# excel-deduction-tool

## 홈택스 카드 매입 세액 공제 불공제 변경 자동화 프로그램

### 홈택스 로그인 기능 지원
**회사명과 ID 및 Password가 담긴 엑셀 파일이 존재해야 함**
- 예시 파일은 아래와 같음
- 별도의 수정이 없는 경우, 두번째 행부터 읽기 시작하며, 회사 코드, 회사 이름, ID, Password 순서로 읽는다.

| 회사 코드 | 회사 이름 | 홈텍스 ID | 홈텍스 Pw |
|-----------|-----------|------------|----------|
| 101       | 가나다(주) | ganadara12 | mabasa45 | 

### 로그인 상태에서 카드 매입 세액 공제 불공제 변경 기능 지원
**공제 불공제 정보가 담긴 엑셀 파일이 존재해야 함**
- 예시 파일은 아래와 같음
- 별도의 수정이 없는 경우, 나머지 컬럼은 필요 없으나, 승인일자, 가맹점사업자번호, 자맹점명, 합계, 공제여부결정 데이터는 존재해야하며, 열의 위치는 같아야함.

| 승인일자    | 카드사           | 카드번호           | 가맹점사업자번호 | 가맹점명                  | 공급가액 | 세액 | 비과세 | 합계 | 가맹점유형 | 업태   | 업종      | 공제여부결정 | 비고        |
|-------------|------------------|--------------------|-----------------|---------------------------|----------|------|--------|------|------------|--------|-----------|--------------|-------------|
| 2000-01-01  | 비씨카드 (주)     | 0000-1111-2222-3333| 111-22-33333    | 가나다라마바사아자 주식회사 | 1,000    | 100  | 0      | 1,100| 법인사업자  | 서비스 | 전자금융업  | 불공제       | 선택불공제   |
