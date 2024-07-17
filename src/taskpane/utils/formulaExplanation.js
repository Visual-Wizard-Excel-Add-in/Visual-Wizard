const FORMULA_EXPLANATION = {
  ABS: "숫자의 절대 값을 반환합니다.\n`ABS(숫자)`",
  ACCRINT:
    "정기적으로 이자를 지급하는 유가 증권의 경과 이자를 반환합니다.\n`ACCRINT(발행일, 결산일, 지급 빈도, 명목 금리, 현재 가치, [지급 빈도], [일 수 기준])`",
  ACCRINTM:
    "만기에 이자를 지급하는 유가 증권의 경과 이자를 반환합니다.\n`ACCRINTM(발행일, 만기일, 명목 금리, 현재 가치, [일 수 기준])`",
  ACOS: "숫자의 아크코사인을 반환합니다.\n`ACOS(숫자)`",
  ACOSH: "숫자의 역 하이퍼볼릭 코사인을 반환합니다.\n`ACOSH(숫자)`",
  ACOT: "숫자의 아크코탄젠트 값을 반환합니다.\n`ACOT(숫자)`",
  ACOTH: "숫자의 하이퍼볼릭 아크코탄젠트 값을 반환합니다.\n`ACOTH(숫자)`",
  ADDRESS:
    "참조를 워크시트의 한 셀에 대한 텍스트로 반환합니다.\n`ADDRESS(행 번호, 열 번호, [참조 유형], [시트 이름])`",
  AGGREGATE:
    "목록 또는 데이터베이스에서 집계 값을 반환합니다.\n`AGGREGATE(함수 번호, 옵션, 참조 1, [참조 2], ...)`",
  AMORDEGRC:
    "감가 상각 계수를 사용하여 매 회계 기간의 감가 상각액을 반환합니다.\n`AMORDEGRC(자산의 초기 비용, 구매일, 첫 사용일, 자산 수명, 회계 기간)\n`",
  AMORLINC:
    "매 회계 기간에 대한 감가 상각액을 반환합니다.\n`AMORLINC(자산의 초기 비용, 구매일, 첫 사용일, 자산 수명, 회계 기간)`",
  AND: "인수가 모두 TRUE이면 TRUE를 반환합니다.\n`AND(조건1, [조건2], ...)`",
  ARABIC: "로마 숫자를 아라비아 숫자로 변환합니다.\n`ARABIC(텍스트)`",
  AREAS: "참조 영역 내의 영역 수를 반환합니다.\n`AREAS(참조)`",
  ARRAYTOTEXT:
    "지정된 범위에서 텍스트 값의 배열을 반환합니다.\n`ARRAYTOTEXT(배열, [형식])`",
  ASC: "문자열에서 영문 전자(더블바이트)나 가타가나 전자를 반자(싱글바이트)로 바꿉니다.\n`ASC(텍스트)`",
  ASIN: "숫자의 아크사인을 반환합니다.\n`ASIN(숫자)`",
  ASINH: "숫자의 역 하이퍼볼릭 사인을 반환합니다.\n`ASINH(숫자)`",
  ATAN: "숫자의 아크탄젠트를 반환합니다.\n`ATAN(숫자)`",
  ATAN2: "x, y 좌표의 아크탄젠트를 반환합니다.\n`ATAN2(x, y)`",
  ATANH: "숫자의 역 하이퍼볼릭 탄젠트를 반환합니다.\n`ATANH(숫자)`",
  AVEDEV:
    "데이터 요소의 절대 편차의 평균을 반환합니다.\n`AVEDEV(숫자1, [숫자2], ...)`",
  AVERAGE: "인수의 평균을 반환합니다.\n`AVERAGE(숫자1, [숫자2], ...)`",
  AVERAGEA:
    "인수의 평균(숫자, 텍스트, 논리값 포함)을 반환합니다.\n`AVERAGEA(값1, [값2], ...)`",
  AVERAGEIF:
    "범위 내에서 주어진 조건에 맞는 모든 셀의 평균(산술 평균)을 반환합니다.\n`AVERAGEIF(범위, 조건, [평균 구할 범위])`",
  AVERAGEIFS:
    "여러 조건에 맞는 모든 셀의 평균(산술 방식)을 반환합니다.\n`AVERAGEIFS(평균 구할 범위, 조건 범위1, 조건1, [조건 범위2, 조건2], ...)`",
  BAHTTEXT:
    "ß(바트) 통화 형식을 사용하여 숫자를 텍스트로 변환합니다.\n`BAHTTEXT(숫자)`",
  BASE: "숫자를 지정된 기수의 텍스트 표현으로 변환합니다.\n`BASE(숫자, 기수, [최소 자릿수])`",
  BESSELI: "수정된 Bessel 함수 In(x) 값을 반환합니다.\n`BESSELI(x, n)`",
  BESSELJ: "Bessel 함수 Jn(x)을 반환합니다.\n`BESSELJ(x, n)`",
  BESSELK: "수정된 Bessel 함수 Kn(x) 값을 반환합니다.\n`BESSELK(x, n)`",
  BESSELY: "Bessel 함수 Yn(x)을 반환합니다.\n`BESSELY(x, n)`",
  BETA_DIST:
    "누적 베타 분포 함수를 반환합니다.\n`BETA_DIST(x, 알파, 베타, [누적], [A], [B])`",
  BETA_INV:
    "지정된 베타 분포에 대한 역 누적 분포 함수를 반환합니다.\n`BETA_INV(확률, 알파, 베타, [A], [B])`",
  BETADIST:
    "누적 베타 분포 함수를 반환합니다.\n`BETADIST(x, 알파, 베타, [A], [B])`",
  BETAINV:
    "지정된 베타 분포에 대한 역 누적 분포 함수를 반환합니다.\n`BETAINV(확률, 알파, 베타, [A], [B])`",
  BIN2DEC: "2진수를 10진수로 변환합니다.\n`BIN2DEC(숫자)`",
  BIN2HEX: "2진수를 16진수로 변환합니다.\n`BIN2HEX(숫자, [자리수])`",
  BIN2OCT: "2진수를 8진수로 변환합니다.\n`BIN2OCT(숫자, [자리수])`",
  BINOM_DIST:
    "개별항 이항 분포 확률을 반환합니다.\n`BINOM_DIST(성공 횟수, 시행 횟수, 성공 확률, 누적)`",
  BINOM_DIST_RANGE:
    "이항 분포를 사용한 시행 결과의 확률을 반환합니다.\n`BINOM_DIST_RANGE(시행 횟수, 성공 확률, 성공 횟수의 하한값, [성공 횟수의 상한값])`",
  BINOM_INV:
    "누적 이항 분포가 기준치 이하가 되는 값 중 최소값을 반환합니다.\n`BINOM_INV(시행 횟수, 성공 확률, 기준값)`",
  BINOMDIST:
    "개별항 이항 분포 확률을 반환합니다.\n`BINOMDIST(성공 횟수, 시행 횟수, 성공 확률, 누적)`",
  BITAND: "두 숫자의 '비트 단위 And'를 반환합니다.\n`BITAND(숫자1, 숫자2)`",
  BITLSHIFT:
    "shift_amount비트씩 왼쪽으로 이동한 값 숫자를 반환합니다.\n`BITLSHIFT(숫자, 이동할 비트 수)`",
  BITOR: "두 숫자의 비트 단위 OR을 반환합니다.\n`BITOR(숫자1, 숫자2)`",
  BITRSHIFT:
    "shift_amount비트씩 오른쪽으로 이동한 값 숫자를 반환합니다.\n`BITRSHIFT(숫자, 이동할 비트 수)`",
  BITXOR:
    "두 숫자의 비트 단위 '배타적 Or'를 반환합니다.\n`BITXOR(숫자1, 숫자2)`",
  BYCOL:
    "각 열에 LAMBDA를 적용하고 결과 배열을 반환합니다.\n`BYCOL(배열, 람다)`",
  BYROW:
    "각 행에 LAMBDA를 적용하고 결과 배열을 반환합니다.\n`BYROW(배열, 람다)`",
  CALL: "DLL(동적 연결 라이브러리) 또는 코드 리소스의 프로시저를 호출합니다.\n`CALL(프로시저, [인자1, 인자2, ...])`",
  CEILING:
    "가장 가까운 정수 또는 가장 가까운 significance의 배수로 숫자를 반올림합니다.\n`CEILING(숫자, significance)`",
  CEILING_MATH:
    "가장 가까운 정수 또는 가장 가까운 significance의 배수로 올림합니다.\n`CEILING_MATH(숫자, [significance], [모드])`",
  CEILING_PRECISE:
    "가장 가까운 정수 또는 가장 가까운 significance의 배수로 내림합니다. 숫자의 부호에 상관없이 숫자는 내림됩니다.\n`CEILING_PRECISE(숫자, [significance])`",
  CELL: "셀의 서식 지정이나 위치, 내용에 대한 정보를 반환합니다.\n`CELL(정보 유형, [참조])`",
  CHAR: "코드 번호에 해당하는 문자를 반환합니다.\n`CHAR(숫자)`",
  CHIDIST:
    "카이 제곱 분포의 단측 검정 확률을 반환합니다.\n`CHIDIST(x, 자유도)`",
  CHIINV:
    "카이 제곱 분포의 역 단측 검정 확률을 반환합니다.\n`CHIINV(확률, 자유도)`",
  CHISQ_DIST:
    "누적 베타 확률 밀도 함수 값을 반환합니다.\n`CHISQ_DIST(x, 자유도, 누적)`",
  CHISQ_DIST_RT:
    "카이 제곱 분포의 단측 검정 확률을 반환합니다.\n`CHISQ_DIST_RT(x, 자유도)`",
  CHISQ_INV:
    "누적 베타 확률 밀도 함수 값을 반환합니다.\n`CHISQ_INV(확률, 자유도)`",
  CHISQ_INV_RT:
    "카이 제곱 분포의 역 단측 검정 확률을 반환합니다.\n`CHISQ_INV_RT(확률, 자유도)`",
  CHISQ_TEST:
    "독립 검증 결과를 반환합니다.\n`CHISQ_TEST(실제 범위, 예상 범위)`",
  CHOOSE:
    "값 목록에서 값을 선택합니다.\n`CHOOSE(인덱스 번호, 값1, [값2], ...)`",
  CHOOSECOLS:
    "배열에서 지정된 열을 반환합니다.\n`CHOOSECOLS(배열, 열 번호1, [열 번호2], ...)`",
  CHOOSEROWS:
    "배열에서 지정된 행을 반환합니다.\n`CHOOSEROWS(배열, 행 번호1, [행 번호2], ...)`",
  CLEAN: "인쇄할 수 없는 모든 문자들을 텍스트에서 제거합니다.\n`CLEAN(텍스트)`",
  CODE: "텍스트 문자열의 첫째 문자를 나타내는 코드값을 반환합니다.\n`CODE(텍스트)`",
  COLUMN: "참조 영역의 열 번호를 반환합니다.\n`COLUMN([참조])`",
  COLUMNS: "참조 영역의 열 수를 반환합니다.\n`COLUMNS(배열)`",
  COMBIN:
    "주어진 개체 수로 만들 수 있는 조합의 수를 반환합니다.\n`COMBIN(숫자, 선택된 숫자)`",
  COMBINA:
    "주어진 개체 수로 만들 수 있는 조합의 수(반복 포함)를 반환합니다.\n`COMBINA(숫자, 선택된 숫자)`",
  COMPLEX:
    "실수부와 허수부의 계수를 복소수로 변환합니다.\n`COMPLEX(실수부, 허수부, [접미사])`",
  CONCAT:
    "여러 범위 및/또는 문자열의 텍스트를 결합하지만 구분 기호나 IgnoreEmpty 인수는 제공하지 않습니다.\n`CONCAT(텍스트1, [텍스트2], ...)`",
  CONCATENATE:
    "여러 텍스트 항목을 한 텍스트 항목으로 조인시킵니다.\n`CONCATENATE(텍스트1, [텍스트2], ...)`",
  CONFIDENCE:
    "모집단 평균의 신뢰 구간을 반환합니다.\n`CONFIDENCE(알파, 표준 편차, 크기)`",
  CONFIDENCE_NORM:
    "모집단 평균의 신뢰 구간을 반환합니다.\n`CONFIDENCE_NORM(알파, 표준 편차, 크기)`",
  CONFIDENCE_T:
    "스튜던트 t-분포를 사용하는 모집단 평균의 신뢰 구간을 반환합니다.\n`CONFIDENCE_T(알파, 표준 편차, 크기)`",
  CONVERT: "다른 단위 체계의 숫자로 변환합니다.\n`CONVERT(숫자, 단위1, 단위2)`",
  CORREL:
    "두 데이터 집합 사이의 상관 계수를 반환합니다.\n`CORREL(배열1, 배열2)`",
  COS: "숫자의 코사인을 반환합니다.\n`COS(숫자)`",
  COSH: "숫자의 하이퍼볼릭 코사인을 반환합니다.\n`COSH(숫자)`",
  COT: "숫자의 하이퍼볼릭 코사인을 반환합니다.\n`COT(숫자)`",
  COTH: "각도의 코탄젠트 값을 반환합니다.\n`COTH(숫자)`",
  COUNT: "인수 목록에서 숫자의 수를 계산합니다.\n`COUNT(값1, [값2], ...)`",
  COUNTA: "인수 목록에서 값의 수를 계산합니다.\n`COUNTA(값1, [값2], ...)`",
  COUNTBLANK: "범위 내에서 빈 셀의 수를 계산합니다.\n`COUNTBLANK(범위)`",
  COUNTIF:
    "범위 내에서 주어진 조건에 맞는 셀의 수를 계산합니다.\n`COUNTIF(범위, 조건)`",
  COUNTIFS:
    "범위 내에서 여러 조건에 맞는 셀의 수를 계산합니다.\n`COUNTIFS(조건 범위1, 조건1, [조건 범위2, 조건2], ...)`",
  COUPDAYBS:
    "이자 지급 기간의 시작일부터 결산일까지의 날짜 수를 반환합니다.\n`COUPDAYBS(결산일, 만기일, 빈도, [일 수 기준])`",
  COUPDAYS:
    "결산일이 들어 있는 이자 지급 기간의 날짜 수를 반환합니다.\n`COUPDAYS(결산일, 만기일, 빈도, [일 수 기준])`",
  COUPDAYSNC:
    "결산일부터 다음 이자 지급일까지의 날짜 수를 반환합니다.\n`COUPDAYSNC(결산일, 만기일, 빈도, [일 수 기준])`",
  COUPNCD:
    "결산일 다음 첫 번째 이자 지급일을 나타내는 숫자를 반환합니다.\n`COUPNCD(결산일, 만기일, 빈도, [일 수 기준])`",
  COUPNUM:
    "결산일과 만기일 사이의 이자 지급 횟수를 반환합니다.\n`COUPNUM(결산일, 만기일, 빈도, [일 수 기준])`",
  COUPPCD:
    "결산일 바로 전 이자 지급일을 나타내는 숫자를 반환합니다.\n`COUPPCD(결산일, 만기일, 빈도, [일 수 기준])`",
  COVAR:
    "각 데이터 요소 쌍에 대한 편차의 곱의 평균(공분산)을 반환합니다.\n`COVAR(배열1, 배열2)`",
  COVARIANCE_P:
    "각 데이터 요소 쌍에 대한 편차의 곱의 평균(공분산)을 반환합니다.\n`COVARIANCE_P(배열1, 배열2)`",
  COVARIANCE_S:
    "두 데이터 집합 사이에서 각 데이터 요소 쌍에 대한 편차의 곱의 평균(표본 집단 공분산)을 반환합니다.\n`COVARIANCE_S(배열1, 배열2)`",
  CRITBINOM:
    "누적 이항 분포가 기준치 이하가 되는 값 중 최소값을 반환합니다.\n`CRITBINOM(시행 횟수, 성공 확률, 기준값)`",
  CSC: "각도의 코시컨트 값을 반환합니다.\n`CSC(숫자)`",
  CSCH: "각도의 하이퍼볼릭 코시컨트 값을 반환합니다.\n`CSCH(숫자)`",
  CUBEKPIMEMBER:
    "KPI(핵심 성과 지표) 이름, 속성 및 측정값을 반환하고 셀에 이름과 속성을 표시합니다.\n`CUBEKPIMEMBER(연결 이름, KPI 이름, KPI 속성, [측정값])`",
  CUBEMEMBER:
    "큐브 계층 구조의 구성원이나 튜플을 반환합니다.\n`CUBEMEMBER(연결 이름, 멤버 표현, [캡션])`",
  CUBEMEMBERPROPERTY:
    "큐브에서 구성원 속성 값을 반환합니다.\n`CUBEMEMBERPROPERTY(연결 이름, 멤버 표현, 속성)`",
  CUBERANKEDMEMBER:
    "집합에서 n번째 또는 순위 내의 구성원을 반환합니다.\n`CUBERANKEDMEMBER(연결 이름, 집합 표현, 인덱스, [캡션])`",
  CUBESET:
    "서버의 큐브에 집합을 만드는 식을 전송하여 계산된 구성원이나 튜플 집합을 정의하고 이 집합을 Microsoft Office Excel에 반환합니다.\n`CUBESET(연결 이름, 집합 표현, [캡션]",
  CUBESETCOUNT: "집합에서 항목 개수를 반환합니다.\n`CUBESETCOUNT(집합)`",
  CUBEVALUE:
    "큐브에서 집계 값을 반환합니다.\n`CUBEVALUE(연결 이름, 멤버 표현1, [멤버 표현2], ...)`",
  CUMIPMT:
    "주어진 기간 중에 납입하는 대출금 이자의 누계액을 반환합니다.\n`CUMIPMT(이자율, 기간, 현재 가치, 시작 기간, 끝 기간, 납입 시점)`",
  CUMPRINC:
    "주어진 기간 중에 납입하는 대출금 원금의 누계액을 반환합니다.\n`CUMPRINC(이자율, 기간, 현재 가치, 시작 기간, 끝 기간, 납입 시점)`",
  DATE: "특정 날짜의 일련 번호를 반환합니다.\n`DATE(연도, 월, 일)`",
  DATEDIF:
    "두 날짜 사이의 일, 월 또는 연도 수를 계산합니다. 이 함수는 경과한 날짜를 계산해야 하는 수식에 유용합니다.\n`DATEDIF(시작일, 종료일, 단위)`",
  DATEVALUE:
    "텍스트 형태의 날짜를 일련 번호로 변환합니다.\n`DATEVALUE(날짜 텍스트)`",
  DAVERAGE:
    "선택한 데이터베이스 항목의 평균을 반환합니다.\n`DAVERAGE(데이터베이스, 필드, 조건)`",
  DAY: "일련 번호를 주어진 달의 날짜로 변환합니다.\n`DAY(일련 번호)`",
  DAYS: "두 날짜 사이의 일 수를 반환합니다.\n`DAYS(종료일, 시작일)`",
  DAYS360:
    "1년을 360일로 하여, 두 날짜 사이의 날짜 수를 계산합니다.\n`DAYS360(시작일, 종료일, [방법])`",
  DB: "정율법을 사용하여 지정한 기간 동안 자산의 감가상각을 반환합니다.\n`DB(자산의 원가, 잔존가치, 자산의 수명, 기간, [월 수])`",
  DBCS: "문자열에서 영문 반자(싱글바이트)나 가타가나 반자를 전자(더블바이트)로 바꿉니다.\n`DBCS(텍스트)`",
  DCOUNT:
    "데이터베이스에서 숫자가 있는 셀의 개수를 계산합니다.\n`DCOUNT(데이터베이스, 필드, 조건)`",
  DCOUNTA:
    "데이터베이스에서 데이터가 들어 있는 셀의 개수를 계산합니다.\n`DCOUNTA(데이터베이스, 필드, 조건)`",
  DDB: "이중 체감법이나 기타 방법을 사용하여 지정된 기간의 감가 상각액을 반환합니다.\n`DDB(자산의 원가, 잔존가치, 자산의 수명, 기간, [요율])`",
  DEC2BIN: "10진수를 2진수로 변환합니다.\n`DEC2BIN(숫자, [자리수])`",
  DEC2HEX: "10진수를 16진수로 변환합니다.\n`DEC2HEX(숫자, [자리수])`",
  DEC2OCT: "10진수를 8진수로 변환합니다.\n`DEC2OCT(숫자, [자리수])`",
  DECIMAL:
    "주어진 기수의 텍스트 표현을 10진수로 변환합니다.\n`DECIMAL(텍스트, 기수)`",
  DEGREES: "라디안 형태의 각도를 도 단위로 바꿉니다.\n`DEGREES(각도)`",
  DELTA: "두 값이 같은지 여부를 검사합니다.\n`DELTA(숫자1, [숫자2])`",
  DEVSQ: "편차의 제곱의 합을 반환합니다.\n`DEVSQ(숫자1, [숫자2], ...)`",
  DGET: "데이터베이스에서 찾을 조건에 맞는 레코드가 하나인 경우 그 레코드를 추출합니다.\n`DGET(데이터베이스, 필드, 조건)`",
  DISC: "유가 증권의 할인율을 반환합니다.\n`DISC(정산일, 만기일, 가격, 환매가, [일수 기준])`",
  DMAX: "선택한 데이터베이스 항목 중에서 최대값을 반환합니다.\n`DMAX(데이터베이스, 필드, 조건)`",
  DMIN: "선택한 데이터베이스 항목 중에서 최소값을 반환합니다.\n`DMIN(데이터베이스, 필드, 조건)`",
  DOLLAR:
    "₩(원) 통화 형식을 사용하여 숫자를 텍스트로 변환합니다.\n`DOLLAR(숫자, [소수 자릿수])`",
  DOLLARDE:
    "분수로 표시된 금액을 소수로 표시된 금액으로 변환합니다.\n`DOLLARDE(분수, 분모)`",
  DOLLARFR:
    "소수로 표시된 금액을 분수로 표시된 금액으로 변환합니다.\n`DOLLARFR(소수, 분모)`",
  DPRODUCT:
    "데이터베이스에서 조건에 맞는 특정 레코드 필드의 값을 곱합니다.\n`DPRODUCT(데이터베이스, 필드, 조건)`",
  DROP: "배열의 시작 또는 끝에서 지정된 개수의 행 또는 열을 제외합니다.\n`DROP(배열, 행 수, [열 수])`",
  DSTDEV:
    "데이터베이스 필드 값들로부터 표본 집단의 표준 편차를 구합니다.\n`DSTDEV(데이터베이스, 필드, 조건)`",
  DSTDEVP:
    "데이터베이스 필드 값들로부터 모집단의 표준 편차를 계산합니다.\n`DSTDEVP(데이터베이스, 필드, 조건)`",
  DSUM: "데이터베이스에서 조건에 맞는 레코드 필드 열 값들의 합을 구합니다.\n`DSUM(데이터베이스, 필드, 조건)`",
  DURATION:
    "정기적으로 이자를 지급하는 유가 증권의 연간 듀레이션을 반환합니다.\n`DURATION(정산일, 만기일, 이율, 수익률, 빈도, [일수 기준])`",
  DVAR: "데이터베이스 필드 값들로부터 표본 집단의 분산을 구합니다.\n`DVAR(데이터베이스, 필드, 조건)`",
  DVARP:
    "데이터베이스 필드 값들로부터 모집단의 분산을 계산합니다.\n`DVARP(데이터베이스, 필드, 조건)`",
  EDATE:
    "지정한 날짜 전이나 후의 개월 수를 나타내는 날짜의 일련 번호를 반환합니다.\n`EDATE(시작일, 개월 수)`",
  EFFECT:
    "연간 실질 이자율을 반환합니다.\n`EFFECT(명목 이자율, 연간 복리 횟수)`",
  ENCODEURL: "URL로 인코딩된 문자열을 반환합니다.\n`ENCODEURL(텍스트)`",
  EOMONTH:
    "지정된 달 수 이전이나 이후 달의 마지막 날의 날짜 일련 번호를 반환합니다.\n`EOMONTH(시작일, 개월 수)`",
  ERF: "오차 함수를 반환합니다.\n`ERF(하한, [상한])`",
  ERF_PRECISE: "오차 함수를 반환합니다.\n`ERF_PRECISE(숫자)`",
  ERFC: "ERF 함수의 여값을 반환합니다.\n`ERFC(숫자)`",
  ERFC_PRECISE:
    "x에서 무한대까지 적분된 ERF 함수의 여값을 반환합니다.\n`ERFC_PRECISE(숫자)`",
  ERROR_TYPE: "오류 유형에 해당하는 숫자를 반환합니다.\n`ERROR_TYPE(오류 값)`",
  EUROCONVERT:
    "숫자를 유로화로, 유로화에서 유로 회원국 통화로 또는 유로화를 매개 통화로 사용하여 숫자를 현재 유로 회원국 통화에서 다른 유로 회원국 통화로 변환(3각 변환)합니다.\n`EUROCONVERT(숫자, 원래 통화, 목표 통화, [유로 고정 비율], [십진수 자리])`",
  EVEN: "가장 가까운 짝수로 숫자를 반올림합니다.\n`EVEN(숫자)`",
  EXACT: "두 텍스트 값이 동일한지 검사합니다.\n`EXACT(텍스트1, 텍스트2)`",
  EXP: "e를 주어진 수만큼 거듭제곱한 값을 반환합니다.\n`EXP(숫자)`",
  EXPAND:
    "지정된 행 및 열 차원으로 배열을 확장하거나 채웁니다.\n`EXPAND(배열, 행 수, 열 수, [채울 값])`",
  EXPON_DIST: "지수 분포값을 반환합니다.\n`EXPON.DIST(x, 람다, 누적)`",
  EXPONDIST: "지수 분포값을 반환합니다.\n`EXPONDIST(x, 람다, 누적)`",
  F_DIST: "F 확률 분포값을 반환합니다.\n`F.DIST(x, 자유도1, 자유도2, 누적)`",
  F_DIST_RT: "F 확률 분포값을 반환합니다.\n`F.DIST.RT(x, 자유도1, 자유도2)`",
  F_INV: "F 확률 분포의 역함수를 반환합니다.\n`F.INV(확률, 자유도1, 자유도2)`",
  F_INV_RT:
    "F 확률 분포의 역함수를 반환합니다.\n`F.INV.RT(확률, 자유도1, 자유도2)`",
  F_TEST: "F-검정 결과를 반환합니다.\n`F.TEST(배열1, 배열2)`",
  FACT: "숫자의 계승값을 반환합니다.\n`FACT(숫자)`",
  FACTDOUBLE: "숫자의 이중 계승값을 반환합니다.\n`FACTDOUBLE(숫자)`",
  FALSE: "논리값 FALSE를 반환합니다.\n`FALSE()`",
  FDIST: "F 확률 분포값을 반환합니다.\n`FDIST(x, 자유도1, 자유도2)`",
  FILTER:
    "사용자가 정의하는 기준에 따라 데이터 범위를 필터링합니다.\n`FILTER(배열, 조건 범위1, [조건 범위2])`",
  FILTERXML:
    "지정된 XPath를 사용하여 XML 콘텐츠의 특정 데이터를 반환합니다.\n`FILTERXML(xml, xpath)`",
  FIND: "텍스트 값에서 다른 텍스트 값을 찾습니다(대/소문자 구분).\n`FIND(찾을 텍스트, 검색할 텍스트, [시작 위치])`",
  FINDB:
    "텍스트 값에서 다른 텍스트 값을 찾습니다(대/소문자 구분).\n`FINDB(찾을 텍스트, 검색할 텍스트, [시작 위치])`",
  FINV: "F 확률 분포의 역함수를 반환합니다.\n`FINV(확률, 자유도1, 자유도2)`",
  FISHER: "피셔 변환 값을 반환합니다.\n`FISHER(x)`",
  FISHERINV: "피셔 변환의 역변환 값을 반환합니다.\n`FISHERINV(y)`",
  FIXED:
    "숫자 표시 형식을 고정 소수점을 사용하는 텍스트로 지정합니다.\n`FIXED(숫자, [소수 자릿수], [쉼표 사용 여부])`",
  FLOOR: "0에 가까워지도록 숫자를 내림합니다.\n`FLOOR(숫자, significance)`",
  FLOOR_MATH:
    "가장 가까운 정수 또는 가장 가까운 significance의 배수로 내림합니다.\n`FLOOR.MATH(숫자, [significance], [모드])`",
  FLOOR_PRECISE:
    "가장 가까운 정수 또는 가장 가까운 significance의 배수로 내림합니다. 숫자의 부호에 상관없이 숫자는 내림됩니다.\n`FLOOR.PRECISE(숫자, [significance])`",
  FORECAST:
    "선형 추세에 따라 값을 반환합니다.\n`FORECAST(x, 알려진 y 값들, 알려진 x 값들)`",
  FORECAST_ETS:
    "AAA 버전의 ETS(지수 평활법) 알고리즘을 사용하여 기존(기록) 값을 기반으로 미래 값을 반환합니다.\n`FORECAST.ETS(대상 날짜, 값들, 타임라인, [시즌 길이], [데이터 완료 여부], [집계 방법])`",
  FORECAST_ETS_CONFINT:
    "지정된 대상 날짜의 예측 값에 대한 신뢰 구간을 반환합니다.\n`FORECAST.ETS.CONFINT(대상 날짜, 값들, 타임라인, [시즌 길이], [데이터 완료 여부], [집계 방법])`",
  FORECAST_ETS_SEASONALITY:
    "Excel에서 지정된 시계열에 대해 감지하는 반복적인 패턴의 길이를 반환합니다.\n`FORECAST.ETS.SEASONALITY(값들, 타임라인, [시즌 길이], [데이터 완료 여부], [집계 방법])`",
  FORECAST_ETS_STAT:
    "시계열 예측의 결과로 통계 값을 반환합니다.\n`FORECAST.ETS.STAT(값들, 타임라인, [시즌 길이], [데이터 완료 여부], [집계 방법], 통계 유형)`",
  FORECAST_LINEAR:
    "기존 값을 기반으로 미래 값을 반환합니다.\n`FORECAST.LINEAR(대상 날짜, 값들, 타임라인)`",
  FORMULATEXT:
    "주어진 참조 영역에 있는 수식을 텍스트로 반환합니다.\n`FORMULATEXT(참조)`",
  FREQUENCY:
    "빈도 분포값을 세로 배열로 반환합니다.\n`FREQUENCY(데이터 배열, 빈도 배열)`",
  FTEST: "F-검정 결과를 반환합니다.\n`FTEST(배열1, 배열2)`",
  FV: "투자의 미래 가치를 반환합니다.\n`FV(이자율, 기간, 납입액, 현재 가치, [납입 시점])`",
  FVSCHEDULE:
    "초기 원금에 일련의 복리 이율을 적용했을 때의 예상 금액을 반환합니다.\n`FVSCHEDULE(원금, 이율 배열)`",
  GAMMA: "감마 함수 값을 반환합니다.\n`GAMMA(숫자)`",
  GAMMA_DIST: "감마 분포값을 반환합니다.\n`GAMMA.DIST(x, 알파, 베타, 누적)`",
  GAMMA_INV:
    "감마 누적 분포의 역함수 값을 반환합니다.\n`GAMMA.INV(확률, 알파, 베타)`",
  GAMMADIST: "감마 분포값을 반환합니다.\n`GAMMADIST(x, 알파, 베타, 누적)`",
  GAMMAINV:
    "감마 누적 분포의 역함수 값을 반환합니다.\n`GAMMAINV(확률, 알파, 베타)`",
  GAMMALN: "감마 함수 Γ(x)의 자연 로그값을 반환합니다.\n`GAMMALN(숫자)`",
  GAMMALN_PRECISE:
    "감마 함수 Γ(x)의 자연 로그값을 반환합니다.\n`GAMMALN.PRECISE(숫자)`",
  GAUSS: "표준 정규 누적 분포값보다 0.5 작은 값을 반환합니다.\n`GAUSS(z)`",
  GCD: "최대 공약수를 반환합니다.\n`GCD(숫자1, [숫자2], ...)`",
  GEOMEAN: "기하 평균을 반환합니다.\n`GEOMEAN(숫자1, [숫자2], ...)`",
  GESTEP: "숫자가 임계값보다 큰지 여부를 검사합니다.\n`GESTEP(숫자, [임계값])`",
  GETPIVOTDATA:
    "피벗 테이블 보고서에 저장된 데이터를 반환합니다.\n`GETPIVOTDATA(데이터 필드, 피벗 테이블 참조, [필드1, 항목1], [필드2, 항목2], ...)`",
  GROWTH:
    "지수 추세를 따라 값을 반환합니다.\n`GROWTH(known_y's, [known_x's], [new_x's], [const])`",
  HARMEAN: "조화 평균을 반환합니다.\n`HARMEAN(숫자1, [숫자2], ...)`",
  HEX2BIN: "16진수를 2진수로 변환합니다.\n`HEX2BIN(숫자, [자리수])`",
  HEX2DEC: "16진수를 10진수로 변환합니다.\n`HEX2DEC(숫자)`",
  HEX2OCT: "16진수를 8진수로 변환합니다.\n`HEX2OCT(숫자, [자리수])`",
  HLOOKUP:
    "배열의 첫 행을 찾아 표시된 셀의 값을 반환합니다.\n`HLOOKUP(검색할 값, 검색할 데이터 포함 범위, 반환 값의 행 번호, [일치 정도])`",
  HOUR: "일련 번호를 시간으로 변환합니다.\n`HOUR(일련 번호)`",
  HSTACK:
    "더 큰 배열을 반환하기 위해 배열을 가로로 순서대로 추가합니다.\n`HSTACK(배열1, [배열2], ...)`",
  HYPERLINK:
    "네트워크 서버, 인트라넷 또는 인터넷에 저장된 문서로 이동할 바로 가기 키를 만듭니다.\n`HYPERLINK(링크 위치, [표시 이름])`",
  HYPGEOM_DIST:
    "초기하 분포값을 반환합니다.\n`HYPGEOM.DIST(성공 횟수, 모집단 크기, 모집단 중 성공 항목 수, 표본 크기, 누적)`",
  HYPGEOMDIST:
    "초기하 분포값을 반환합니다.\n`HYPGEOMDIST(성공 횟수, 모집단 크기, 모집단 중 성공 항목 수, 표본 크기, 누적)`",
  IF: "입력한 로직의 참, 거짓 여부에 따라 설정한 반환 값을 반환합니다.\n`IF(논리식, 참일 때 반환할 값, 거짓일 때 반환할 값)`",
  IFERROR:
    "수식이 오류이면 사용자가 지정한 값을 반환하고, 그렇지 않으면 수식 결과를 반환합니다.\n`IFERROR(값, 오류일 경우 반환할 값)`",
  IFNA: "식이 #N/A로 계산되면 지정한 값을 반환하고, 그렇지 않으면 식의 결과를 반환합니다.\n`IFNA(값, #N/A일 경우 반환할 값)`",
  IFS: "하나 이상의 조건이 충족되는지 여부를 확인하고 첫 번째 TRUE 정의에 해당하는 값을 반환합니다.\n`IFS(조건1, 참일 때 반환할 값1, 조건2, 참일 때 반환할 값2, ...)`",
  IMABS: "복소수의 절대값을 반환합니다.\n`IMABS(복소수)`",
  IMAGE:
    "지정된 원본에서 이미지를 반환합니다.\n`IMAGE(소스, [대체 텍스트], [너비], [높이], [유지 비율])`",
  IMAGINARY: "복소수의 허수부 계수를 반환합니다.\n`IMAGINARY(복소수)`",
  IMARGUMENT:
    "각도가 라디안으로 표시되는 테타 인수를 반환합니다.\n`IMARGUMENT(복소수)`",
  IMCONJUGATE: "복소수의 켤레 복소수를 반환합니다.\n`IMCONJUGATE(복소수)`",
  IMCOS: "복소수의 코사인을 반환합니다.\n`IMCOS(복소수)`",
  IMCOSH: "복소수의 하이퍼볼릭 코사인 값을 반환합니다.\n`IMCOSH(복소수)`",
  IMCOT: "복소수의 코탄젠트 값을 반환합니다.\n`IMCOT(복소수)`",
  IMCSC: "복소수의 코시컨트 값을 반환합니다.\n`IMCSC(복소수)`",
  IMCSCH: "복소수의 하이퍼볼릭 코시컨트 값을 반환합니다.\n`IMCSCH(복소수)`",
  IMDIV: "두 복소수의 나눗셈 몫을 반환합니다.\n`IMDIV(복소수1, 복소수2)`",
  IMEXP: "복소수의 지수를 반환합니다.\n`IMEXP(복소수)`",
  IMLN: "복소수의 자연 로그값을 반환합니다.\n`IMLN(복소수)`",
  IMLOG10: "복소수의 밑이 10인 로그값을 반환합니다.\n`IMLOG10(복소수)`",
  IMLOG2: "복소수의 밑이 2인 로그값을 반환합니다.\n`IMLOG2(복소수)`",
  IMPOWER: "복소수의 멱을 반환합니다.\n`IMPOWER(복소수, 지수)`",
  IMPRODUCT: "복소수의 곱을 반환합니다.\n`IMPRODUCT(복소수1, 복소수2, ...)`",
  IMREAL: "복소수의 실수부 계수를 반환합니다.\n`IMREAL(복소수)`",
  IMSEC: "복소수의 시컨트 값을 반환합니다.\n`IMSEC(복소수)`",
  IMSECH: "복소수의 하이퍼볼릭 시컨트 값을 반환합니다.\n`IMSECH(복소수)`",
  IMSIN: "복소수의 사인을 반환합니다.\n`IMSIN(복소수)`",
  IMSINH: "복소수의 하이퍼볼릭 사인 값을 반환합니다.\n`IMSINH(복소수)`",
  IMSQRT: "복소수의 제곱근을 반환합니다.\n`IMSQRT(복소수)`",
  IMSUB: "두 복소수 간의 차를 반환합니다.\n`IMSUB(복소수1, 복소수2)`",
  IMSUM: "복소수의 합을 반환합니다.\n`IMSUM(복소수1, 복소수2, ...)`",
  IMTAN: "복소수의 탄젠트 값을 반환합니다.\n`IMTAN(복소수)`",
  INDEX:
    "인덱스를 사용하여 참조 영역이나 배열의 값을 선택합니다.\n`INDEX(배열, 행 번호, [열 번호], [영역 번호])`",
  INDIRECT:
    "텍스트 값으로 표시된 참조를 반환합니다.\n`INDIRECT(참조 텍스트, [A1])`",
  INFO: "현재 운영 환경에 대한 정보를 반환합니다.\n`INFO(정보 유형)`",
  INT: "가장 가까운 정수로 숫자를 내림합니다.\n`INT(숫자)`",
  INTERCEPT:
    "선형 회귀선의 절편을 반환합니다.\n`INTERCEPT(known_y's, known_x's)`",
  INTRATE:
    "완전 투자한 유가 증권의 이자율을 반환합니다.\n`INTRATE(정산일, 만기일, 투자액, 수령액, [일수 기준])`",
  IPMT: "일정 기간 동안의 투자 금액에 대한 이자 지급액을 반환합니다.\n`IPMT(이자율, 기간, 전체 기간 수, 현재 가치, [미래 가치], [납입 시점])`",
  IRR: "일련의 현금 흐름에 대한 내부 수익률을 반환합니다.\n`IRR(현금 흐름, [추정 값])`",
  ISBLANK: "값이 비어 있으면 TRUE를 반환합니다.\n`ISBLANK(값)`",
  ISERR: "값이 #N/A를 제외한 오류 값이면 TRUE를 반환합니다.\n`ISERR(값)`",
  ISERROR: "값이 오류 값이면 TRUE를 반환합니다.\n`ISERROR(값)`",
  ISEVEN: "숫자가 짝수이면 TRUE를 반환합니다.\n`ISEVEN(숫자)`",
  ISFORMULA:
    "수식을 포함하는 셀에 대한 참조가 있으면 TRUE를 반환합니다.\n`ISFORMULA(참조)`",
  ISLOGICAL: "값이 논리값이면 TRUE를 반환합니다.\n`ISLOGICAL(값)`",
  ISNA: "값이 #N/A 오류 값이면 TRUE를 반환합니다.\n`ISNA(값)`",
  ISNONTEXT: "값이 텍스트가 아니면 TRUE를 반환합니다.\n`ISNONTEXT(값)`",
  ISNUMBER: "값이 숫자이면 TRUE를 반환합니다.\n`ISNUMBER(값)`",
  ISO_CEILING:
    "가장 가까운 정수 또는 significance의 배수로 반올림한 숫자를 반환합니다.\n`ISO.CEILING(숫자, [significance])`",
  ISODD: "숫자가 홀수이면 TRUE를 반환합니다.\n`ISODD(숫자)`",
  ISOMITTED:
    "LAMBDA의 값이 누락되었는지 확인하고 TRUE 또는 FALSE를 반환합니다.\n`ISOMITTED(값)`",
  ISOWEEKNUM:
    "지정된 날짜에 따른 해당 연도의 ISO 주 번호를 반환합니다.\n`ISOWEEKNUM(날짜)`",
  ISPMT:
    "일정 기간 동안의 투자에 대한 이자 지급액을 계산합니다.\n`ISPMT(이자율, 기간, 전체 기간 수, 대출 금액)`",
  ISREF: "값이 셀 주소이면 TRUE를 반환합니다.\n`ISREF(값)`",
  ISTEXT: "값이 텍스트이면 TRUE를 반환합니다.\n`ISTEXT(값)`",
  JIS: "문자열에서 반자(싱글바이트) 문자를 전자(더블바이트)로 바꿉니다.\n`JIS(텍스트)`",
  KURT: "데이터 집합의 첨도를 반환합니다.\n`KURT(배열)`",
  LAMBDA:
    "재사용 가능한 사용자 지정 함수를 만들고 이름을 사용하여 호출합니다.\n`LAMBDA(파라미터, 계산할 수식)`",
  LARGE: "데이터 집합에서 k번째로 큰 값을 반환합니다.\n`LARGE(배열, k)`",
  LCM: "최소 공배수를 반환합니다.\n`LCM(숫자1, [숫자2], ...)`",
  LEFT: "텍스트 값에서 맨 왼쪽의 문자를 반환합니다.\n`LEFT(텍스트, [문자 수])`",
  LEFTB:
    "텍스트 값에서 맨 왼쪽의 문자를 반환합니다.\n`LEFTB(텍스트, [문자 수])`",
  LEN: "텍스트 문자열의 문자 수를 반환합니다.\n`LEN(텍스트)`",
  LENB: "텍스트 문자열의 바이트 수를 반환합니다.\n`LENB(텍스트)`",
  LET: "계산 결과에 이름을 지정합니다.\n`LET(이름1, 이름값1, 수식/계산)`",
  LINEST:
    "선형 추세의 매개 변수를 반환합니다.\n`LINEST(known_y's, [known_x's], [const], [stats])`",
  LN: "숫자의 자연 로그를 반환합니다.\n`LN(숫자)`",
  LOG: "지정한 밑에 대한 로그를 반환합니다.\n`LOG(숫자, [밑])`",
  LOG10: "밑이 10인 로그값을 반환합니다.\n`LOG10(숫자)`",
  LOGEST:
    "지수 추세의 매개 변수를 반환합니다.\n`LOGEST(known_y's, [known_x's], [const], [stats])`",
  LOGINV:
    "로그 정규 누적 분포의 역함수 값을 반환합니다.\n`LOGINV(확률, 평균, 표준편차)`",
  LOGNORM_DIST:
    "로그 정규 누적 분포값을 반환합니다.\n`LOGNORM.DIST(x, 평균, 표준편차, 누적)`",
  LOGNORM_INV:
    "로그 정규 누적 분포의 역함수 값을 반환합니다.\n`LOGNORM.INV(확률, 평균, 표준편차)`",
  LOGNORMDIST:
    "로그 정규 누적 분포값을 반환합니다.\n`LOGNORMDIST(x, 평균, 표준편차, 누적)`",
  LOOKUP:
    "벡터나 배열에서 값을 찾습니다.\n`LOOKUP(검색값, 검색벡터, 결과벡터)`",
  LOWER: "텍스트를 소문자로 변환합니다.\n`LOWER(텍스트)`",
  MAKEARRAY:
    "LAMBDA를 적용하여 지정된 행 및 열 크기의 계산된 배열을 반환합니다.\n`MAKEARRAY(행수, 열수, LAMBDA)`",
  MAP: "새 값을 생성하기 위해 LAMBDA를 적용하여 배열의 각 값을 새 값에 매핑하여 형성된 배열을 반환합니다.\n`MAP(배열, LAMBDA)`",
  MATCH:
    "참조 영역이나 배열에서 값을 찾습니다.\n`MATCH(검색값, 검색범위, [일치유형])`",
  MAX: "인수 목록에서 최대값을 반환합니다.\n`MAX(숫자1, [숫자2], ...)`",
  MAXA: "인수 목록에서 최대값(숫자, 텍스트, 논리값 포함)을 반환합니다.\n`MAXA(값1, [값2], ...)`",
  MAXIFS:
    "주어진 조건에 맞는 셀에서 최대값을 반환합니다.\n`MAXIFS(최대값 범위, 조건 범위1, 조건1, [조건 범위2, 조건2], ...)`",
  MDETERM: "배열의 행렬식을 반환합니다.\n`MDETERM(배열)`",
  MDURATION:
    "가정된 액면가 $100에 대한 유가 증권의 수정된 Macauley 듀레이션을 반환합니다.\n`MDURATION(정산일, 만기일, 이율, 수익률, 빈도, [일수 기준])`",
  MEDIAN: "주어진 수들의 중간값을 반환합니다.\n`MEDIAN(숫자1, [숫자2], ...)`",
  MID: "지정된 위치에서 시작하여 특정 개수의 문자를 텍스트 문자열에서 반환합니다.\n`MID(텍스트, 시작 위치, 문자 수)`",
  MIDB: "지정된 위치에서 시작하여 특정 개수의 문자를 텍스트 문자열에서 반환합니다.\n`MIDB(텍스트, 시작 위치, 문자 수)`",
  MIN: "인수 목록에서 최소값을 반환합니다.\n`MIN(숫자1, [숫자2], ...)`",
  MINA: "인수 목록에서 최소값(숫자, 텍스트, 논리값 포함)을 반환합니다.\n`MINA(값1, [값2], ...)`",
  MINIFS:
    "주어진 조건에 맞는 셀에서 최소값을 반환합니다.\n`MINIFS(최소값 범위, 조건 범위1, 조건1, [조건 범위2, 조건2], ...)`",
  MINUTE: "일련 번호를 분으로 변환합니다.\n`MINUTE(일련 번호)`",
  MINVERSE: "배열의 역행렬을 반환합니다.\n`MINVERSE(배열)`",
  MIRR: "다른 이율로 형성되는 양의 현금 흐름과 음의 현금 흐름에 대한 내부 수익률을 반환합니다.\n`MIRR(현금흐름, 재투자율, 자본비용)`",
  MMULT: "두 배열의 행렬 곱을 반환합니다.\n`MMULT(배열1, 배열2)`",
  MOD: "나눗셈의 나머지를 반환합니다.\n`MOD(숫자, 나눗수)`",
  MODE: "데이터 집합에서 가장 일반적인 값을 반환합니다.\n`MODE(숫자1, [숫자2], ...)`",
  MODE_MULT:
    "배열이나 데이터 범위에서 가장 자주 발생하는 값의 세로 배열을 반환합니다.\n`MODE.MULT(숫자1, [숫자2], ...)`",
  MODE_SNGL:
    "데이터 집합에서 가장 많이 나오는 값을 반환합니다.\n`MODE.SNGL(숫자1, [숫자2], ...)`",
  MONTH: "일련 번호를 월로 변환합니다.\n`MONTH(일련 번호)`",
  MROUND: "원하는 배수로 반올림된 수를 반환합니다.\n`MROUND(숫자, 배수)`",
  MULTINOMIAL:
    "각 계승값의 곱에 대한 합계의 계승값 비율을 반환합니다.\n`MULTINOMIAL(숫자1, [숫자2], ...)`",
  MUNIT: "지정된 차원에 대한 단위 행렬을 반환합니다.\n`MUNIT(차원)`",
  N: "숫자로 변환된 값을 반환합니다.\n`N(값)`",
  NA: "#N/A 오류 값을 반환합니다.\n`NA()`",
  NEGBINOM_DIST:
    "음 이항 분포값을 반환합니다.\n`NEGBINOM.DIST(성공횟수, 실패횟수, 성공확률, 누적)`",
  NEGBINOMDIST:
    "음 이항 분포값을 반환합니다.\n`NEGBINOMDIST(성공횟수, 실패횟수, 성공확률, 누적)`",
  NETWORKDAYS:
    "두 날짜 사이의 전체 작업 일수를 반환합니다.\n`NETWORKDAYS(시작일, 종료일, [휴일])`",
  NETWORKDAYS_INTL:
    "주말인 날짜와 해당 날짜 수를 나타내는 매개 변수를 사용하여 두 날짜 사이의 전체 작업일 수를 반환합니다.\n`NETWORKDAYS.INTL(시작일, 종료일, [주말], [휴일])`",
  NOMINAL:
    "명목상의 연이율을 반환합니다.\n`NOMINAL(실제 연이율, 연간 복리 횟수)`",
  NORM_DIST:
    "정규 누적 분포값을 반환합니다.\n`NORM.DIST(x, 평균, 표준편차, 누적)`",
  NORM_INV:
    "정규 누적 분포의 역함수 값을 반환합니다.\n`NORM.INV(확률, 평균, 표준편차)`",
  NORM_S_DIST: "표준 정규 누적 분포값을 반환합니다.\n`NORM.S.DIST(z, 누적)`",
  NORM_S_INV:
    "표준 정규 누적 분포의 역함수 값을 반환합니다.\n`NORM.S.INV(확률)`",
  NORMDIST:
    "정규 누적 분포값을 반환합니다.\n`NORMDIST(x, 평균, 표준편차, 누적)`",
  NORMINV:
    "정규 누적 분포의 역함수 값을 반환합니다.\n`NORMINV(확률, 평균, 표준편차)`",
  NORMSDIST: "표준 정규 누적 분포값을 반환합니다.\n`NORMSDIST(z)`",
  NORMSINV: "표준 정규 누적 분포의 역함수 값을 반환합니다.\n`NORMSINV(확률)`",
  NOT: "인수의 논리 역을 반환합니다.\n`NOT(논리값)`",
  NOW: "현재 날짜 및 시간의 일련 번호를 반환합니다.\n`NOW()`",
  NPER: "투자의 기간을 반환합니다.\n`NPER(이자율, 납입액, 현재가치, [미래가치], [납입시점])`",
  NPV: "주기적인 현금 흐름과 할인율을 기준으로 투자의 순 현재 가치를 반환합니다.\n`NPV(할인율, 현금흐름1, [현금흐름2], ...)`",
  NUMBERVALUE:
    "로캘에 영향을 받지 않으면서 텍스트를 숫자로 변환합니다.\n`NUMBERVALUE(텍스트, [소수점 구분 기호], [천 단위 구분 기호])`",
  OCT2BIN: "8진수를 2진수로 변환합니다.\n`OCT2BIN(숫자, [자리수])`",
  OCT2DEC: "8진수를 10진수로 변환합니다.\n`OCT2DEC(숫자)`",
  OCT2HEX: "8진수를 16진수로 변환합니다.\n`OCT2HEX(숫자, [자리수])`",
  ODD: "가장 가까운 홀수로 숫자를 반올림합니다.\n`ODD(숫자)`",
  ODDFPRICE:
    "첫 이수 기간이 경상 이수 기간과 다른 유가 증권의 액면가 $100당 가격을 반환합니다.\n`ODDFPRICE(정산일, 만기일, 이율, 수익률, 액면가, 정기 이자 지급일, [일수 기준])`",
  ODDFYIELD:
    "첫 이수 기간이 경상 이수 기간과 다른 유가 증권의 수익률을 반환합니다.\n`ODDFYIELD(정산일, 만기일, 이율, 액면가, 정기 이자 지급일, 가격, [일수 기준])`",
  ODDLPRICE:
    "마지막 이수 기간이 경상 이수 기간과 다른 유가 증권의 액면가 $100당 가격을 반환합니다.\n`ODDLPRICE(정산일, 만기일, 이율, 수익률, 액면가, 정기 이자 지급일, [일수 기준])`",
  ODDLYIELD:
    "마지막 이수 기간이 경상 이수 기간과 다른 유가 증권의 수익률을 반환합니다.\n`ODDLYIELD(정산일, 만기일, 이율, 가격, 액면가, 정기 이자 지급일, [일수 기준])`",
  OFFSET:
    "주어진 참조 영역으로부터의 참조 영역 간격을 반환합니다.\n`OFFSET(참조, 행 이동 수, 열 이동 수, [높이], [너비])`",
  OR: "인수가 하나라도 TRUE이면 TRUE를 반환합니다.\n`OR(논리값1, [논리값2], ...)`",
  PDURATION:
    "투자 금액이 지정된 값에 도달할 때까지 필요한 기간을 반환합니다.\n`PDURATION(이자율, 현재가치, 미래가치)`",
  PEARSON: "피어슨 곱 모멘트 상관 계수를 반환합니다.\n`PEARSON(배열1, 배열2)`",
  PERCENTILE: "범위에서 k번째 백분위수를 반환합니다.\n`PERCENTILE(배열, k)`",
  PERCENTILE_EXC:
    "범위에서 k번째 백분위수를 반환합니다. 이때 k는 경계값을 제외한 0에서 1 사이의 수입니다.\n`PERCENTILE.EXC(배열, k)`",
  PERCENTILE_INC:
    "범위에서 k번째 백분위수를 반환합니다.\n`PERCENTILE.INC(배열, k)`",
  PERCENTRANK:
    "데이터 집합의 값에 대한 백분율 순위를 반환합니다.\n`PERCENTRANK(배열, x, [유의미 자릿수])`",
  PERCENTRANK_EXC:
    "데이터 집합에서 경계값을 제외한 0에서 1 사이의 백분율 순위를 반환합니다.\n`PERCENTRANK.EXC(배열, x, [유의미 자릿수])`",
  PERCENTRANK_INC:
    "데이터 집합의 값에 대한 백분율 순위를 반환합니다.\n`PERCENTRANK.INC(배열, x, [유의미 자릿수])`",
  PERMUT:
    "주어진 개체 수로 만들 수 있는 순열의 수를 반환합니다.\n`PERMUT(숫자, 선택된 숫자)`",
  PERMUTATIONA:
    "전체 개체에서 선택하여 주어진 개체 수(반복 포함)로 만들 수 있는 순열의 수를 반환합니다.\n`PERMUTATIONA(숫자, 선택된 숫자)`",
  PHI: "표준 정규 분포의 밀도 함수 값을 반환합니다.\n`PHI(x)`",
  PHONETIC: "텍스트 문자열에서 윗주 문자를 추출합니다.\n`PHONETIC(텍스트)`",
  PI: "원주율(파이) 값을 반환합니다.\n`PI()`",
  PMT: "연금의 정기 납입액을 반환합니다.\n`PMT(이자율, 기간, 현재가치, [미래가치], [납입시점])`",
  POISSON: "포아송 확률 분포값을 반환합니다.\n`POISSON(x, 평균, 누적)`",
  POISSON_DIST:
    "포아송 확률 분포값을 반환합니다.\n`POISSON.DIST(x, 평균, 누적)`",
  POWER:
    "밑수를 지정한 만큼 거듭제곱한 결과를 반환합니다.\n`POWER(숫자, 지수)`",
  PPMT: "일정 기간 동안의 투자에 대한 원금의 지급액을 반환합니다.\n`PPMT(이자율, 기간, 전체 기간 수, 현재 가치, [미래 가치], [납입 시점])`",
  PRICE:
    "정기적으로 이자를 지급하는 유가 증권의 액면가 $100당 가격을 반환합니다.\n`PRICE(정산일, 만기일, 이율, 수익률, 정기이자, [일수기준])`",
  PRICEDISC:
    "할인된 유가 증권의 액면가 $100당 가격을 반환합니다.\n`PRICEDISC(정산일, 만기일, 할인율, 액면가, [일수기준])`",
  PRICEMAT:
    "만기일에 이자를 지급하는 유가 증권의 액면가 $100당 가격을 반환합니다.\n`PRICEMAT(정산일, 만기일, 이율, 수익률, 액면가, [일수기준])`",
  PROB: "영역 내의 값이 두 한계값 사이에 있을 확률을 반환합니다.\n`PROB(배열, 확률, 하한값, [상한값])`",
  PRODUCT: "인수를 곱합니다.\n`PRODUCT(숫자1, [숫자2], ...)`",
  PROPER:
    "텍스트 값에 있는 각 단어의 첫째 문자를 대문자로 바꿉니다.\n`PROPER(텍스트)`",
  PV: "투자의 현재 가치를 반환합니다.\n`PV(이자율, 기간, 납입액, [미래가치], [납입시점])`",
  QUARTILE:
    "데이터 집합에서 사분위수를 반환합니다.\n`QUARTILE(배열, 사분위수)`",
  QUARTILE_EXC:
    "데이터 집합에서 경계값을 제외한 0에서 1 사이의 사분위수를 반환합니다.\n`QUARTILE.EXC(배열, 사분위수)`",
  QUARTILE_INC:
    "데이터 집합에서 사분위수를 반환합니다.\n`QUARTILE.INC(배열, 사분위수)`",
  QUOTIENT: "나눗셈 몫의 정수 부분을 반환합니다.\n`QUOTIENT(숫자, 나눗수)`",
  RADIANS: "도 단위로 표시된 각도를 라디안으로 변환합니다.\n`RADIANS(각도)`",
  RAND: "0과 1 사이의 난수를 반환합니다.\n`RAND()`",
  RANDARRAY:
    "0과 1 사이의 임의의 숫자 배열을 반환합니다. 단, 채울 행과 열 수, 최소값과 최대값 및 정수 또는 소수값을 반환할지 여부를 지정할 수 있습니다.\n`RANDARRAY(행 수, 열 수, [최소값], [최대값], [정수 여부])`",
  RANDBETWEEN:
    "지정한 두 수 사이의 난수를 반환합니다.\n`RANDBETWEEN(최소값, 최대값)`",
  RANK: "숫자 목록 내에서 지정한 숫자의 크기 순위를 반환합니다.\n`RANK(숫자, 참조, [순위 방식])`",
  RANK_AVG:
    "수 목록 내에서 지정한 수의 크기 순위를 반환합니다.\n`RANK.AVG(숫자, 참조, [순위 방식])`",
  RANK_EQ:
    "수 목록 내에서 지정한 수의 크기 순위를 반환합니다.\n`RANK.EQ(숫자, 참조, [순위 방식])`",
  RATE: "연금의 기간별 이자율을 반환합니다.\n`RATE(기간, 납입액, 현재가치, 미래가치, [납입시점], [추정치])`",
  RECEIVED:
    "완전 투자 유가 증권에 대해 만기 시 수령하는 금액을 반환합니다.\n`RECEIVED(정산일, 만기일, 투자액, 이율, [일수 기준])`",
  REDUCE:
    "각 값에 LAMBDA를 적용하고 누산기에서 총 값을 반환하여 배열을 누산 값으로 줄입니다.\n`REDUCE(초기값, 배열, LAMBDA)`",
  REGISTER_ID:
    "지정한 DLL(동적 연결 라이브러리) 또는 코드 리소스의 레지스터 ID를 반환합니다.\n`REGISTER.ID(DLL, 프로시저, [인수1, 인수2, ...])`",
  REPLACE:
    "텍스트 내의 문자를 바꿉니다.\n`REPLACE(텍스트, 시작 위치, 바꿀 문자 수, 새 텍스트)`",
  REPLACEB:
    "텍스트 내의 문자를 바꿉니다.\n`REPLACEB(텍스트, 시작 위치, 바꿀 바이트 수, 새 텍스트)`",
  REPT: "텍스트를 지정된 횟수만큼 반복합니다.\n`REPT(텍스트, 횟수)`",
  RIGHT:
    "텍스트 값에서 맨 오른쪽의 문자를 반환합니다.\n`RIGHT(텍스트, [문자 수])`",
  RIGHTB:
    "텍스트 값에서 맨 오른쪽의 문자를 반환합니다.\n`RIGHTB(텍스트, [바이트 수])`",
  ROMAN:
    "아라비아 숫자를 텍스트인 로마 숫자로 변환합니다.\n`ROMAN(숫자, [형식])`",
  ROUND: "수를 지정한 자릿수로 반올림합니다.\n`ROUND(숫자, 자릿수)`",
  ROUNDDOWN: "0에 가까워지도록 수를 내림합니다.\n`ROUNDDOWN(숫자, 자릿수)`",
  ROUNDUP: "0에서 멀어지도록 수를 올림합니다.\n`ROUNDUP(숫자, 자릿수)`",
  ROW: "참조의 행 번호를 반환합니다.\n`ROW([참조])`",
  ROWS: "참조 영역에 있는 행 수를 반환합니다.\n`ROWS(배열)`",
  RRI: "투자 수익에 해당하는 이자율을 반환합니다.\n`RRI(기간, 현재가치, 미래가치)`",
  RSQ: "피어슨 곱 모멘트 상관 계수의 제곱을 반환합니다.\n`RSQ(배열1, 배열2)`",
  RTD: "COM 자동화를 지원하는 프로그램으로부터 실시간 데이터를 가져옵니다.\n`RTD(프로그램, 서버, 주제1, [주제2], ...)`",
  SCAN: "각 값에 LAMBDA를 적용하여 배열을 검사하고 각 중간 값이 있는 배열을 반환합니다.\n`SCAN(초기값, 배열, LAMBDA)`",
  SEARCH:
    "지정한 텍스트 값을 다른 텍스트 값 내에서 찾습니다(대/소문자 구분 안 함).\n`SEARCH(찾을 텍스트, 검색할 텍스트, [시작 위치])`",
  SEARCHB:
    "지정한 텍스트 값을 다른 텍스트 값 내에서 찾습니다(대/소문자 구분 안 함).\n`SEARCHB(찾을 텍스트, 검색할 텍스트, [시작 위치])`",
  SEC: "각도의 시컨트 값을 반환합니다.\n`SEC(숫자)`",
  SECH: "각도의 하이퍼볼릭 시컨트 값을 반환합니다.\n`SECH(숫자)`",
  SECOND: "일련 번호를 초로 변환합니다.\n`SECOND(일련 번호)`",
  SEQUENCE:
    "1, 2, 3, 4와 같은 배열에서 일련 번호 목록을 생성합니다.\n`SEQUENCE(행 수, 열 수, 시작값, 증가량)`",
  SERIESSUM:
    "수식에 따라 멱급수의 합을 반환합니다.\n`SERIESSUM(x, n, m, 계수)`",
  SHEET: "참조된 시트의 시트 번호를 반환합니다.\n`SHEET(참조)`",
  SHEETS: "참조 영역에 있는 시트 수를 반환합니다.\n`SHEETS(참조)`",
  SIGN: "수의 부호값을 반환합니다.\n`SIGN(숫자)`",
  SIN: "지정된 각도의 사인을 반환합니다.\n`SIN(각도)`",
  SINH: "숫자의 하이퍼볼릭 사인을 반환합니다.\n`SINH(숫자)`",
  SKEW: "분포의 왜곡도를 반환합니다.\n`SKEW(배열)`",
  SKEW_P:
    "왜곡도란 평균에 대한 분포의 비대칭 정도를 나타냅니다.\n`SKEW.P(배열)`",
  SLN: "한 기간 동안 정액법에 의한 자산의 감가 상각액을 반환합니다.\n`SLN(원가, 잔존가치, 수명)`",
  SLOPE: "선형 회귀선의 기울기를 반환합니다.\n`SLOPE(known_y's, known_x's)`",
  SMALL: "데이터 집합에서 k번째로 작은 값을 반환합니다.\n`SMALL(배열, k)`",
  SORT: "범위 또는 배열의 내용을 정렬합니다.\n`SORT(배열, [정렬 인덱스], [정렬 순서], [정렬 기준])`",
  SORTBY:
    "대응되는 범위 또는 배열의 값을 기준으로 범위 또는 배열의 내용을 정렬합니다.\n`SORTBY(배열, 정렬 기준 범위1, [정렬 순서1], [정렬 기준 범위2], [정렬 순서2], ...)`",
  SQRT: "양의 제곱근을 반환합니다.\n`SQRT(숫자)`",
  SQRTPI: "(number * pi)의 제곱근을 반환합니다.\n`SQRTPI(숫자)`",
  STANDARDIZE: "정규화된 값을 반환합니다.\n`STANDARDIZE(x, 평균, 표준편차)`",
  STDEV: "표본 집단의 표준 편차를 구합니다.\n`STDEV(숫자1, [숫자2], ...)`",
  STDEV_P: "모집단의 표준 편차를 계산합니다.\n`STDEV.P(숫자1, [숫자2], ...)`",
  STDEV_S: "표본 집단의 표준 편차를 구합니다.\n`STDEV.S(숫자1, [숫자2], ...)`",
  STDEVA:
    "표본 집단의 표준 편차(숫자, 텍스트, 논리값 포함)를 구합니다.\n`STDEVA(값1, [값2], ...)`",
  STDEVP:
    "전체 모집단의 표준 편차를 계산합니다.\n`STDEVP(숫자1, [숫자2], ...)`",
  STDEVPA:
    "모집단의 표준 편차(숫자, 텍스트, 논리값 포함)를 계산합니다.\n`STDEVPA(값1, [값2], ...)`",
  STEYX:
    "회귀분석에 의해 예측한 y값의 표준 오차를 각 x값에 대하여 반환합니다.\n`STEYX(known_y's, known_x's)`",
  STOCKHISTORY:
    "금융 상품에 대한 기록 데이터를 검색합니다.\n`STOCKHISTORY(종목코드, 시작날짜, 종료날짜, [속성], [빈도], [빈도매개변수], [수정종가])`",
  SUBSTITUTE:
    "텍스트 문자열에서 기존 텍스트를 새 텍스트로 바꿉니다.\n`SUBSTITUTE(텍스트, 기존 텍스트, 새 텍스트, [바꿀 횟수])`",
  SUBTOTAL:
    "목록이나 데이터베이스의 부분합을 반환합니다.\n`SUBTOTAL(함수 번호, 참조1, [참조2], ...)`",
  SUM: "인수의 합을 구합니다.\n`SUM(숫자1, [숫자2], ...)`",
  SUMIF:
    "주어진 조건에 의해 지정된 셀들의 합을 구합니다.\n`SUMIF(범위, 조건, [합할 범위])`",
  SUMIFS:
    "범위 내에서 여러 조건에 맞는 셀들의 합을 구합니다.\n`SUMIFS(합할 범위, 조건 범위1, 조건1, [조건 범위2, 조건2], ...)`",
  SUMPRODUCT:
    "배열의 대응되는 구성 요소끼리 곱해서 그 합을 반환합니다.\n`SUMPRODUCT(배열1, [배열2], ...)`",
  SUMSQ: "인수의 제곱의 합을 반환합니다.\n`SUMSQ(숫자1, [숫자2], ...)`",
  SUMX2MY2:
    "두 배열에서 대응값의 제곱을 구한 다음 그 차이의 합을 반환합니다.\n`SUMX2MY2(배열_x, 배열_y)`",
  SUMX2PY2:
    "두 배열에서 대응값의 제곱을 구한 다음 그 합의 합을 반환합니다.\n`SUMX2PY2(배열_x, 배열_y)`",
  SUMXMY2:
    "두 배열에서 대응값의 차이를 구한 다음 그 제곱의 합을 반환합니다.\n`SUMXMY2(배열_x, 배열_y)`",
  SWITCH:
    "값의 목록에 대한 식을 계산하고 첫 번째 일치하는 값에 해당하는 결과를 반환합니다. 일치하는 항목이 없는 경우 선택적 기본값이 반환될 수 있습니다.\n`SWITCH(식, 값1, 결과1, [값2, 결과2], ..., [기본값])`",
  SYD: "지정된 감가 상각 기간 중 자산의 감가 상각액을 연수 합계법으로 반환합니다.\n`SYD(원가, 잔존가치, 수명, 기간)`",
  T: "인수를 텍스트로 변환합니다.\n`T(값)`",
  T_DIST:
    "스튜던트 t-분포의 백분율(확률값)을 반환합니다.\n`T.DIST(x, 자유도, 누적)`",
  T_DIST_2T:
    "스튜던트 t-분포의 백분율(확률값)을 반환합니다.\n`T.DIST.2T(x, 자유도)`",
  T_DIST_RT: "스튜던트 t-분포값을 반환합니다.\n`T.DIST.RT(x, 자유도)`",
  T_INV:
    "스튜던트 t-분포의 t-값을 확률과 자유도에 대한 함수로 반환합니다.\n`T.INV(확률, 자유도)`",
  T_INV_2T:
    "스튜던트 t-분포의 역함수 값을 반환합니다.\n`T.INV.2T(확률, 자유도)`",
  T_TEST:
    "스튜던트 t-검정에 근거한 확률을 반환합니다.\n`T.TEST(배열1, 배열2, 꼬리수, 종류)`",
  TAKE: "배열의 시작 또는 끝에서 지정된 수의 연속 행 또는 열을 반환합니다.\n`TAKE(배열, 행 수, [열 수])`",
  TAN: "숫자의 탄젠트를 반환합니다.\n`TAN(숫자)`",
  TANH: "숫자의 하이퍼볼릭 탄젠트를 반환합니다.\n`TANH(숫자)`",
  TBILLEQ:
    "국채에 대해 채권에 해당하는 수익률을 반환합니다.\n`TBILLEQ(정산일, 만기일, 할인율)`",
  TBILLPRICE:
    "국채에 대해 액면가 $100당 가격을 반환합니다.\n`TBILLPRICE(정산일, 만기일, 할인율)`",
  TBILLYIELD: "국채의 수익률을 반환합니다.\n`TBILLYIELD(정산일, 만기일, 가격)`",
  TDIST: "스튜던트 t-분포값을 반환합니다.\n`TDIST(x, 자유도, 꼬리수)`",
  TEXT: "숫자 표시 형식을 지정하고 텍스트로 변환합니다.\n`TEXT(값, 형식 텍스트)`",
  TEXTAFTER:
    "주어진 문자 또는 문자열 다음에 나오는 텍스트를 반환합니다.\n`TEXTAFTER(텍스트, 구분자, [인스턴스 번호], [일치 모드], [검색 모드])`",
  TEXTBEFORE:
    "주어진 문자 또는 문자열 앞에 나오는 텍스트를 반환합니다.\n`TEXTBEFORE(텍스트, 구분자, [인스턴스 번호], [일치 모드], [검색 모드])`",
  TEXTJOIN:
    "여러 범위 및/또는 문자열의 텍스트를 결합합니다.\n`TEXTJOIN(구분자, 빈 셀 무시, 텍스트1, [텍스트2], ...)`",
  TEXTSPLIT:
    "열 및 행 구분 기호를 사용하여 텍스트 문자열을 분할합니다.\n`TEXTSPLIT(텍스트, 열 구분자, [행 구분자], [열 제한], [행 제한])`",
  TIME: "특정 시간의 일련 번호를 반환합니다.\n`TIME(시, 분, 초)`",
  TIMEVALUE:
    "텍스트 형태의 시간을 일련 번호로 변환합니다.\n`TIMEVALUE(시간 텍스트)`",
  TINV: "학생 t-분포의 역함수 값을 반환합니다.\n`TINV(확률, 자유도)`",
  TOCOL: "단일 열의 배열을 반환합니다.\n`TOCOL(배열, [무시할 값])`",
  TODAY: "오늘 날짜의 일련 번호를 반환합니다.\n`TODAY()`",
  TOROW: "단일 행의 배열을 반환합니다.\n`TOROW(배열, [무시할 값])`",
  TRANSPOSE: "배열의 행과 열을 바꿉니다.\n`TRANSPOSE(배열)`",
  TREND:
    "선형 추세에 따라 값을 반환합니다.\n`TREND(known_y's, [known_x's], [new_x's], [const])`",
  TRIM: "텍스트에서 공백을 제거합니다.\n`TRIM(텍스트)`",
  TRIMMEAN:
    "데이터 집합의 양 끝값을 제외하고 평균을 구합니다.\n`TRIMMEAN(배열, 제외할 비율)`",
  TRUE: "논리값 TRUE를 반환합니다.\n`TRUE()`",
  TRUNC: "수의 소수점 이하를 버립니다.\n`TRUNC(숫자, [자릿수])`",
  TTEST:
    "스튜던트 t-검정에 근거한 확률을 반환합니다.\n`TTEST(배열1, 배열2, 꼬리수, 종류)`",
  TYPE: "값의 데이터 형식을 나타내는 숫자를 반환합니다.\n`TYPE(값)`",
  UNICHAR:
    "주어진 숫자 값이 참조하는 유니코드 문자를 반환합니다.\n`UNICHAR(숫자)`",
  UNICODE:
    "텍스트의 첫 문자에 해당하는 숫자(코드 포인트)를 반환합니다.\n`UNICODE(텍스트)`",
  UNIQUE:
    "목록 또는 범위에서 고유 값의 목록을 반환합니다.\n`UNIQUE(배열, [일치 기준], [일치 모드])`",
  UPPER: "텍스트를 대문자로 변환합니다.\n`UPPER(텍스트)`",
  VALUE: "텍스트 인수를 숫자로 변환합니다.\n`VALUE(텍스트)`",
  VALUETOTEXT: "지정된 값의 텍스트를 반환합니다.\n`VALUETOTEXT(값, [형식])`",
  VAR: "표본의 분산을 구합니다.\n`VAR(숫자1, [숫자2], ...)`",
  VAR_P: "모집단의 분산을 계산합니다.\n`VAR.P(숫자1, [숫자2], ...)`",
  VAR_S: "표본 집단의 분산을 구합니다.\n`VAR.S(숫자1, [숫자2], ...)`",
  VARA: "표본 집합의 분산(숫자, 텍스트, 논리값 포함)을 구합니다.\n`VARA(값1, [값2], ...)`",
  VARP: "모집단의 분산을 계산합니다.\n`VARP(숫자1, [숫자2], ...)`",
  VARPA:
    "모집단의 분산(숫자, 텍스트, 논리값 포함)을 계산합니다.\n`VARPA(값1, [값2], ...)`",
  VDB: "일정 또는 일부 기간 동안 체감법으로 자산의 감가 상각액을 반환합니다.\n`VDB(원가, 잔존가치, 수명, 시작 기간, 종료 기간, [요율], [월 수])`",
  VLOOKUP:
    "배열의 첫째 열을 찾아 행 쪽으로 이동하여 셀 값을 반환합니다.\n`VLOOKUP(검색할 값, 검색할 데이터 포함 범위, 반환할 열 번호, [일치 정도])`",
  VSTACK:
    "배열을 세로 및 순서대로 추가하여 더 큰 배열을 반환합니다.\n`VSTACK(배열1, [배열2], ...)`",
  WEBSERVICE: "웹 서비스에서 데이터를 반환합니다.\n`WEBSERVICE(url)`",
  WEEKDAY: "일련 번호를 요일로 변환합니다.\n`WEEKDAY(일련 번호, [반환 형식])`",
  WEEKNUM:
    "일련 번호를 해당 주가 일 년 중 몇 번째 주인지 나타내는 숫자로 변환합니다.\n`WEEKNUM(일련 번호, [시작일])`",
  WEIBULL: "와이블 분포값을 반환합니다.\n`WEIBULL(x, 알파, 베타, 누적)`",
  WEIBULL_DIST:
    "와이블 분포값을 반환합니다.\n`WEIBULL.DIST(x, 알파, 베타, 누적)`",
  WORKDAY:
    "특정 일(시작 날짜)의 전이나 후의 날짜 수에서 주말이나 휴일을 제외한 날짜 수, 즉 평일 수를 반환합니다.\n`WORKDAY(시작일, 일 수, [휴일])`",
  WORKDAY_INTL:
    "주말인 날짜와 해당 날짜 수를 나타내는 매개 변수를 사용하여 지정된 작업 일수 이전 또는 이후 날짜의 일련 번호를 반환합니다.\n`WORKDAY.INTL(시작일, 일 수, [주말], [휴일])`",
  WRAPCOLS:
    "지정된 수의 요소 이후에 제공된 행 또는 값 열을 열로 래핑합니다.\n`WRAPCOLS(배열, 열 크기, [채울 값])`",
  WRAPROWS:
    "지정된 수의 요소 뒤에 행별로 제공된 행 또는 값 열을 래핑합니다.\n`WRAPROWS(배열, 행 크기, [채울 값])`",
  XIRR: "비정기적일 수도 있는 현금 흐름의 내부 회수율을 반환합니다.\n`XIRR(현금흐름, 날짜, [추정치])`",
  XLOOKUP:
    "범위 또는 배열을 검색하고 검색된 첫 번째 일치 항목에 해당하는 항목을 반환합니다. 일치 항목이 없는 경우 XLOOKUP 함수는 가장 가까운(대략적인) 일치 항목을 반환할 수 있습니다.\n`XLOOKUP(검색 값, 검색 배열, 반환 배열, [일치 없음], [일치 모드], [검색 모드])`",
  XMATCH:
    "배열이나 셀 범위에서 항목의 상대적 위치를 반환합니다.\n`XMATCH(검색 값, 검색 배열, [일치 모드], [검색 모드])`",
  XNPV: "비정기적일 수도 있는 현금 흐름의 순 현재 가치를 반환합니다.\n`XNPV(이율, 현금흐름, 날짜)`",
  XOR: "모든 인수의 논리 배타적 OR을 반환합니다.\n`XOR(논리값1, [논리값2], ...)`",
  YEAR: "일련 번호를 연도로 변환합니다.\n`YEAR(일련 번호)`",
  YEARFRAC:
    "start_date와 end_date 사이의 날짜 수가 일 년 중 차지하는 비율을 반환합니다.\n`YEARFRAC(시작일, 종료일, [일수 기준])`",
  YIELD:
    "정기적으로 이자를 지급하는 유가 증권의 수익률을 반환합니다.\n`YIELD(정산일, 만기일, 이율, 가격, 상환금, 빈도, [일수 기준])`",
  YIELDDISC:
    "국채와 같이 할인된 유가 증권의 연 수익률을 반환합니다.\n`YIELDDISC(정산일, 만기일, 가격, 상환금, [일수 기준])`",
  YIELDMAT:
    "만기 시 이자를 지급하는 유가 증권의 연 수익률을 반환합니다.\n`YIELDMAT(정산일, 만기일, 이율, 가격, 상환금, [일수 기준])`",
  Z_TEST: "z-test의 편측 확률값을 추출합니다.\n`Z.TEST(배열, x, [sigma])`",
  ZTEST: "z-test의 단측 확률값을 추출합니다.\n`ZTEST(배열, x, [sigma])`",
};

export default FORMULA_EXPLANATION;
