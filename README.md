```mermaid
flowchart LR
    A[Account Info] -- 계좌번호.xlsx --> B[Python Script] --> C[Dashboard]
```


# 계좌 정보
> 삼성증권에서 계좌별 거래내역 받을때 파일명은 **_계좌번호.xlsx_** 이다.

## 계좌 번호
|`account_number`|`account_name`|`description`|
|---|---|---|
|7145601085-01|종합계좌|부동산 매매 비용|
|7145601085-14|ISA|만기후 연금계좌 이전|
|7145601085-25|금현물||
|7162157033-01|자사주|자사주 받는 계좌|
|7146770750-29|IRP||
|7187285150-28|퇴직연금(DC)|
|7146826853-15|연금계좌(동부)|동부에서 연금보험이전한 금액, 안세공|
|7147729613-15|연금계좌(회사)|회사지원금으로 납입|
|7158765677-15|연금계좌(월급)|내월급에서 추가 납입|

## 계좌 거래내역 파일 정보
계좌별 거래내역 엑셀 파일을 읽고 해석하기위한 정보는 [TransactionHistory](./TransactionHistory.md) 에서 확인할 수 있다.


# 자산 보유 현황 
> 자산을 종목별로 현황을 보여줄때 필요한 정보 

|item|kor|description|
|--|--|--|
|`account_name`|계좌명||
|`account_number`|계좌번호||
|`financial_product_name`|종목명||
|`ticker`|티커||
|`currency`|통화||
|`number_of_shares`|보유수량||
|`avg_purchase_price`|매수평단가|`total_cost / number_of_shares`|
|`market_price`|현재가|현재 1주(최소 구매단위) 평가액|
|`total_cost`|매수총금액|현재 주식을 사기위해 지불한 총금액|
|`total_market_value`|평가총금액|`number_of_shares * market_price`|
|`cumulative_realized_gain`|누적실현손익|`total_gain - total_loss`|
|`unrealized_gain`|미실현손익|`total_market_value - total_cost`|
|`unrealized_roi`|미실현수익률|`(unrealized_gain / total_cost) * 100`|
|`total_return`|총수익|`cumulative_realized_gain + unrealized_gain`|

## 티커 획득법

### 코스피/코드닥
```py
import FinanceDataReader as fdr

def get_krx_ticker(name: str) -> str | None:
    # KRX 전체 상장 종목 목록 가져오기
    df = fdr.StockListing('KRX')
    
    # 종목명(Name)으로 필터링하여 종목코드(Code) 반환
    result = df[df['Name'] == name]
    if not result.empty:
        return result.iloc[0]['Code']
    return None

# 실행 예시
print(get_krx_ticker("삼성전자"))  # 005930
print(get_krx_ticker("NAVER"))     # 035420
```

### 미국주식
```py
import yfinance as yf

def get_us_ticker(company_name: str) -> str | None:
    search = yf.Search(company_name, max_results=5)
    quotes = search.quotes
    
    if quotes:
        # 가장 유사도가 높은 첫 번째 주식 티커 반환
        return quotes[0].get('symbol')
    return None

# 실행 예시
print(get_us_ticker("Apple"))      # AAPL
print(get_us_ticker("Microsoft"))  # MSFT
print(get_us_ticker("Nvidia"))     # NVDA
```



# 엑셀 읽기
```bash
pip install pandas openpyxl
```

```py
import pandas as pd

# 엑셀 파일 읽기
file_path = "sample.xlsx"
df = pd.read_excel(file_path)

# 방법 A: 각 행을 딕셔너리 형태로 순회 (열 이름으로 접근하기 편함)
for index, row in df.iterrows():
    print(f"[{index}행] {row.to_dict()}")
    # 특정 컬럼 값 접근 예시: row['이름'], row['나이']

# 방법 B: 튜플 형태로 순회 (대용량 데이터 시 처리 속도가 훨씬 빠름)
# for row in df.itertuples(index=True):
#     print(row)
```