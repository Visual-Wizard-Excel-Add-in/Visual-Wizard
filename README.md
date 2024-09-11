# <img alt="Visual Wizard logo 32x32" src="https://github.com/user-attachments/assets/7310a6b4-0d8a-4b98-8b2a-d034050d1aec" width="32"/> Visual Wizard

<div align="center">
  <img alt="Visual Wizard logo 150x150" src="https://github.com/user-attachments/assets/7310a6b4-0d8a-4b98-8b2a-d034050d1aec" width="150"/>

<b>Visual Wizard</b>는 Excel 사용자들을 위해 개발한 간편한 Add-in으로서
<br/>
<b>입력된 수식의 분석, 서식 저장, 매크로 녹화, 시트 추출 등</b>의 기능을 제공합니다.

<b>Visual Wizard</b>와 함께 생산성을 높여보세요!

</div>

<br/>

## 🔗 링크

<div align="center">

[**Visual Wizard Repository**](https://github.com/Visual-Wizard-Excel-Add-in/Visual-Wizard)
| [**Local 사용 설명서**](https://fair-gram-629.notion.site/95ba0f0286234a58a9a381b6940364a4?pvs=4)

</div>

<br/>

## 🗂️ 목차

  - [**🌼 소개**](#소개)
  - [**⚒️ 기술 스택**](#️기술-스택)
    - [Client](#client)
    - [State Management](#state-management)
    - [Test](#test)
    - [Deployment](#deployment)
  - [**🗓️ 개발 일정**](#️-개발-일정)
  - [**🔍 기능 소개**](#기능-소개)
    - [**1. 수식 탭**](#1-수식-탭)
      - 1-1. 정보 기능
      - 1-2. 참조 기능
      - 1-3. 순서 기능
    - [**2. 서식 탭**](#2-서식-탭)
      - 2-1. 셀 서식 기능
      - 2-2. 차트 서식 기능
    - [**3. 매크로 탭**](#3-매크로-탭)
      - 3-1. 매크로 녹화 기능
      - 3-2. 매크로 설정 기능
    - [**4. 유효성 탭**](#4-유효성-탭)
      - 4-1. 유효성 검사 기능
      - 4-2. 수식 테스트 기능
    - [**5. 공유하기 탭**](#5-공유하기-탭)
      - 5-1. 추출하기 기능
  - [**🔥 기술적 과제**](#-기술적-과제)
    - [**1. Excel과 Add-in과의 통신은 어떻게 이루어질까?**](#1-excel과-add-in과의-통신은-어떻게-이루어질까)
      - 1-1. Add-in이란?
      - 1-2. Excel과 Add-in의 데이터 교류 방법
    - [**2. Office JS와 비동기성**](#2-office-js와-비동기성)
      - 2-1. Excel 데이터 조회 및 관리
      - 2-2. 비동기 작업의 관리
  - [**🚨 기획 변경**](#-기획-변경)
    - [**1. 매크로 녹화 기능 구현 이슈**](#1-매크로-녹화-기능-구현-이슈)
    - [**2. 추출하기 기능 구현 이슈**](#2-추출하기-기능-구현-이슈)
  - [**⌛️ 회고**](#️-회고)

<br/>

## 🌼 소개

Visual Wizard는 사용자가 효율적으로 엑셀 파일을 파악하고 관리할 수 있게 해주는 Excel Add-in입니다.

**처음 접하는 Excel 파일이 어떻게 동작하는지 파악하는 데 힘이 들었던 적이 있으신가요?
여기저기 흩어져 있는 편의 기능을 찾느라 기능 탭을 전부 열어보신 적이 있나요?**

이런 불편을 해결하고자 Visual Wizard는 입력된 **수식에 대한 정보와 관계된 셀**들을 쉽게 찾아볼 수 있습니다.

또한 **자주 사용하는 편의 기능들을 Add-in에 모아놓아** 헤매지 않고 원하는 기능을 바로 사용하실 수 있습니다.

Visual Wizard와 함께 업무의 불필요한 수고를 줄여보세요!

<br/>

## ⚒️ 기술 스택

### Client

![Javascript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=JavaScript&logoColor=white) ![React](https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB) ![OFFICEJSAPI](https://img.shields.io/badge/OFFICEJSAPI-ffa416.svg?style=for-the-badge&logo=&logoColor=white) ![FLUENTUI](https://img.shields.io/badge/FLUENTUI-lightgray?style=for-the-badge&logo=FLUENTUI&logoColor=white)
![Tailwind](https://img.shields.io/badge/tailwindCSS-06B6D4?style=for-the-badge&logo=tailwindcss&logoColor=white)

### State Management

![ZUSTAND](https://img.shields.io/badge/ZUSTAND-lightgray?style=for-the-badge&logo=ZUSTAND&logoColor=blue)

### Test

![Vitest](https://img.shields.io/badge/Vitest-%2344A833.svg?style=for-the-badge&logo=vitest&logoColor=white)

### Deployment

![NETLIFY](https://img.shields.io/badge/Netlify-00C7B7?style=for-the-badge&logo=netlify&logoColor=white)

<br/>

## 🗓️ 개발 일정

<div align="center">

### 프로젝트 기간: 2024.07.08 ~ 2024.07.31

<details>
  <summary>세부 일정</summary>

  <table>
    <tr>
      <th>주차</th>
        <th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;내용</th>
      </tr>
      <tr>
        <td>1주차</td>
        <td>
          - 아이디어 수집, 선정<br/>
          - 기술 스택 결정 및 학습<br/>
          - KANBAN 작성
        </td>
      </tr>
      <tr>
        <td>2주차</td>
        <td>
          - 프로젝트 세팅<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- Client: React, Office JS API, Fluent UI, Tailwind CSS<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- ESLint, Prettier, Husky 설정<br/><br/>
          - Add-in 정적 UI 구현<br/><br/>
          - 수식 탭 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- 정보 기능 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 선택한 셀 정보 불러오기 함수 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 함수 설명 파일 추가<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- 참조 기능 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 셀 서식 저장 및 적용 함수 작성<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 셀 강조 및 복원 함수 작성<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- 순서 기능 구현<br>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 선택한 셀 수식 판별 함수 작성<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 함수별 인수 및 조건 판별 함수 작성<br/>
        </td>
      </tr>
      <tr>
        <td>3주차</td>
        <td>
          - 프리셋 목록 추가 및 불러오기 함수 구현&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/><br/>
          - 서식 탭 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- 셀 서식 기능 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 자료 구조화 및 Office 저장소 저장<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 선택한 프리셋에 서식 저장 및 불러오기 함수 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;-차트 서식 기능 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-차트 서식 저장 및 불러오기 함수 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-자료 구조화 및 Office 저장소 저장<br/><br/>
          - 매크로 탭 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;- 매크로 녹화 기능 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 시트 수정, 차트 추가, 표 추가 및 수정 내역 기록 함수 구현<br/>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 각 액션 감지 이벤트에 함수 등록<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 녹화된 액션 재생 함수 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;- 매크로 설정 기능 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 매크로 녹화 내역 목록화 표시 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 기록된 셀 값, 표 범위, 차트 타입 사용자 임의 변경 함수 구현<br/><br/>
         - 유효성 탭 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;- 에러 셀 강조 기능 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 사용중인 셀에서 에러 검출 및 강조 함수 구현<br/>
        </td>
      </tr>
      <tr>
        <td>4주차</td>
        <td>
         - 유효성 탭 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;- 사용중인 가장 바깥 셀 기능 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;- 수식 테스트 기능 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 선택한 셀의 수식이 참조하는 인수 목록화 표시<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 수식의 인수 변경 함수 구현<br/>
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 변경한 수식 계산 함수 구현<br/><br/>
       - 공유하기 탭 구현<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;- 추출하기 기능 구현<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 사용자 선택 추출 범위 VBA 전달 용 시트 생성 함수 구현<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- VBA 트리거 용 시트 생성 함수 구현<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 트리거 시트 생성 감지 및 사용자 선택 추출 범위 인식 VBA 함수 구현<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 각 범위 별 추출 및 새로운 통합 문서 생성 VBA 함수 구현
        </td>
      </tr>
    </table>
  </div>
</details>

<br/>

## 🔍 기능 소개

### 1. 수식 탭

<details>
  <summary>
   <strong>1-1. 정보 기능</strong>
  </summary>
  <div align="center">
    <img width="200" alt="정보 기능" src="https://github.com/user-attachments/assets/ffbc3701-5633-4166-90fd-b3e73d3c09f2">
  </div>
  
- 선택한 셀의 값이 수식일 경우, 수식에 포함된 함수들의 정보를 보여줍니다.
- 하단에는 Excel 함수 사전 사이트로 이동할 수 있는 링크가 첨부돼 있습니다.

  참조:
  <a href="https://support.microsoft.com/ko-kr/office/excel-함수-사전순-b3944572-255d-4efb-bb96-c6d90033e188#bm19">Excel 함수 사전</a>
</details>

<details>
  <summary>
    <strong>1-2 참조 기능</strong>
  </summary>
  <div align="center">
    <img width="700" alt="참조 기능" src="https://github.com/user-attachments/assets/6415e5cb-7144-4701-aa67-8c97821086cc">
  </div>
  
- 선택한 셀이 수식일 경우, 수식이 참조하는 인수와 결과 셀의 주소 및 값을 표시합니다.
- 토글 버튼 클릭 시, 인수 및 결과 셀의 스타일을 변경하여 강조됩니다.
- 다시 클릭할 경우 강조가 해제됩니다.
</details>

<details>
  <summary>
    <strong>1-3 순서 기능</strong>
  </summary>
  <div align="center">
    <img width="700" alt="순서 기능" src="https://github.com/user-attachments/assets/b7445ee5-f9d4-4792-a02e-88bba3342de2">
  </div>
  
- 선택한 셀이 수식일 경우, 해당 수식에 속한 함수들을 순서대로 나열합니다.
- 각 함수명 클릭 시, 함수의 인수, 조건, 결과값을 보여줍니다.
</details>

<br/>

### 2. 서식 탭

<details>
  <summary>
    <strong>2-1. 셀 서식 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "셀 서식 기능" src = "https://github.com/user-attachments/assets/e96dca26-f075-48c2-9012-48ee887734f0">
  </div>
  
- 설정해 놓은 셀 서식을 프리셋에 저장하거나, 저장한 서식을 다른 셀에 적용할 수 있습니다.
</details>

<details>
  <summary>
    <strong>2-2. 차트 서식 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "차트 서식 기능" src = "https://github.com/user-attachments/assets/1f8a5231-240c-4cb7-9bce-3c1ba4df3a78">
  </div>
  
- 만들어 놓은 차트의 서식을 프리셋에 저장하거나, 저장한 서식을 차트에 적용할 수 있습니다.
</details>

<br/>

### 3. 매크로 탭

<details>
  <summary>
    <strong>3-1. 매크로 녹화 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "매크로 녹화 기능" src = "https://github.com/user-attachments/assets/6b93424c-dc6a-4126-9991-3cd5e1872a25">
  </div>
  
- 녹화 버튼 클릭 후, 버튼을 다시 클릭할 때까지의 `셀 입력`, `셀 서식 변경`, `차트 추가`, `표 추가` 액션을 녹화합니다.
</details>

<details>
  <summary>
    <strong>3-2. 매크로 설정 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "매크로 설정 기능" src = "https://github.com/user-attachments/assets/37c47542-69b1-4d02-9d9d-734a6aaa8d90">
  </div>
  
- 기록한 매크로 내역을 수정할 수 있습니다.
- 수정할 수 있는 액션의 타입은 `셀 내용 변경`, `차트 추가`, `표 추가`입니다.
</details>

<br/>

### 4. 유효성 탭

<details>
  <summary>
    <strong>4-1. 유효성 검사 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "매크로 녹화 기능" src = "https://github.com/user-attachments/assets/1106c5a3-1a31-49ec-860c-d848c23cff8d">
  </div>
  
- 토글 할 경우, 에러가 있는 셀들을 강조합니다.
- 값이 입력된 가장 마지막 위치의 셀 주소를 표시합니다.
- 마지막 셀이 의도한 셀 주소와 상이할 경우 Excel 파일의 용량이 비정상적으로 증가하는 것을 방지할 수 있습니다.
</details>

<details>
  <summary>
    <strong>4-2. 수식 테스트 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "매크로 녹화 기능" src = "https://github.com/user-attachments/assets/3fe0155c-623d-42a6-b882-951929ceef59">
  </div>

- 선택한 셀의 수식을 확인하고, 임의로 인수 및 조건을 수정할 수 있습니다.
- 수식이 원하는 기댓값을 출력하는지 테스트할 수 있습니다.
</details>

<br/>

### 5. 공유하기 탭

<details>
  <summary>
    <strong>5-1. 추출하기 기능</strong>
  </summary>
  <div align = "center">
    <img width = "700" alt = "매크로 녹화 기능" src = "https://github.com/user-attachments/assets/8723690e-bdba-4d40-88d1-2d7e812b2f3f">
  </div>

  - 작성한 Excel 문서의 특정 부분을 새로운 통합 문서로 추출할 수 있습니다.
  - `선택한 범위`, `현재 시트` 중 선택할 수 있습니다.
  </details>
  
  <br/>
  
  ## 🔥 기술적 과제
  
  ### **1. Excel과 Add-in과의 통신은 어떻게 이루어질까?**
  
  <details>
    <summary>
      <strong>1-1. Add-in이란?</strong>
    </summary>
    
  Add-in이란 Excel과 같은 Office Application 내에서 다양한 추가기능을 제공하는 Web 기반 확장프로그램입니다.
  
  - Excel '홈 탭'의 가장 우측에서 접근이 가능합니다.
  
  <div align="center">
    <img width="600" alt="excel home tab" src="https://github.com/user-attachments/assets/0f87ad23-7071-43b3-8395-eed92b062181">
  </div>
  
  - 클릭 시 Excel 화면 우측에 웹 기반 패널이 열립니다.
  
  <div align="center">
    <img width="600" alt="excel with add-in" src="https://github.com/user-attachments/assets/fefc909c-0c24-4eed-9a7f-1c524da1e5f6">
  </div>
  
  해당 사이드 패널은 웹 페이지 이므로 `JavaScript`, `HTML`, `CSS`를 통한 개발이 가능합니다.
</details>

<details>
  <summary>
    <strong>1-2. Excel과 Add-in의 데이터 교류 방법</strong>
  </summary>
  
  Web과 Excel 간의 상호작용을 위해 **Office JavaScript API**(Office JS)를 활용해야 했습니다.
  
  해당 Office JS는 성능 최적화를 위해 통합 문서의 모든 내용을 한 번에 불러오지 않습니다.
  이를 위해 필요한 데이터만 선택적으로 로드하고 동기화하는 방식을 사용합니다.
  
  - **ex) 선택한 셀의 주소와 수식을 로드하기.**
  
  <div align="center">
    <table>
      <tr>
        <td>
          <img width="400" alt="getSelectedCell(code)" src="https://github.com/user-attachments/assets/a106ab89-f4da-4ba3-a9f9-0d74babaf1f5" />
        </td>
        <td>
<img width="400" alt="getSelectedCell(image)" src="https://github.com/user-attachments/assets/1ab4767b-9d3e-420e-9e00-86f132a23c4e">
        </td>
      </tr>
    </table>
  </div>
  
  <br/>
  
  - **ex) `A1`셀의 값 변경하기.**
  
  <div align="center">
    <table>
      <tr>
        <td>
          <img width="400" alt="changeA1(code)" src="https://github.com/user-attachments/assets/922dded7-556f-4c3c-8905-1d077c8ea0fa" />
        </td>
        <td>
          <img width="400" alt="changeA1(image)" src="https://github.com/user-attachments/assets/a56cb44c-9acc-412b-8349-8540d1817ef7" />
        </td>
      </tr>
    </table>
  </div>
  
  이런 식으로 Excel과 Add-in과의 데이터 교환은 Add-in에서의 요청 작업을 모아 **context.sync()** 로 요청한 작업을 수행하고 동기화하는 과정을 통해서 이루어집니다.
  
  데이터 접근은 API 상에서 정해진 Worksheet, Range, Chart 등의 객체에 대해서만 접근할 수 있으며, 모든 작업은 비동기적으로 이루어집니다.
  
  이러한 특징을 바탕으로 필요한 객체들의 속성과 메서드를 효율적으로 활용하기 위하여
  <a href="https://learn.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview">Excel JavaScript API 공식 문서</a>를 참고하는데 가장 많은 시간을 할애했습니다.
</details>

<br/>

### 2. Office JS와 비동기성

<details>
  <summary>
    <strong>2-1. Excel 데이터 조회 및 관리</strong>
  </summary>
  
  Add-in 개발에 있어서 가장 신경 쓴 점은 Office JS에 적응하는 것이었습니다.<br/>
  Excel 데이터에 접근하기 위해선 Office JS의 함수와 메서드에 익숙해져야 했으며, 그 모두가 비동기적으로 이루어진다는 환경에 적응해야 했습니다.
  
  Excel에는 다양한 종류의 객체들이 존재하며, 해당 객체들의 속성과 지원하는 메서드가 모두 다르기 때문에 하나의 속성을 조회하고 수정하기 위해선, console과 공식 문서를 끊임없이 봐야 했습니다.
    
  만약 특정 셀의 서식을 저장하고 싶다면 해당 셀의 서식을 지칭하는 속성들의 종류를 조회하고, 해당 속성들이 지원하는 값들을 조사해야 했습니다.
  
  - 공식 문서에서 지원 속성 조회
    <div align="center">
      <img width="600" alt="Office JS document" src="https://github.com/user-attachments/assets/f31f1c73-e322-470e-9814-bfa1b85d2452">
    </div>
  - 실제 속성 구조 확인
  
  <table>
    <tr>
      <td>
        <img width = "600" alt="cell attribute dev console 1" src = "https://github.com/user-attachments/assets/bf936358-9059-4a38-a2d7-43b59abf2d62">
      </td>
      <td>
      ➡
      </td>
      <td>
        <img width = "600" alt="cell attribute dev console 2" src = "https://github.com/user-attachments/assets/3c866e4b-51b1-4c54-b41b-a77c14a3e755">
      </td>
    </tr>
  </table>
  
  <br/>
  
  - 로드한 속성들
  
  <table>
    <tr>
      <td>
        <img width="600" alt="load cell (code)" src = "https://github.com/user-attachments/assets/96f681f0-c00a-4869-a72e-f8337d8f277c">
      </td>
      <td>
        <img width="600" alt="get borders (code)" src = "https://github.com/user-attachments/assets/a52072b8-b7cc-4727-a616-1b7ae3ce8bd2">
      </td>
    </tr>
  </table>
  
  위와 같이 명시적으로 서식 속성들을 `load()` 한 후 `context.sync()`를 통해 동기화한 후에야 해당 속성들을 사용할 수 있습니다.
  
  해당 서식 속성들을 효율적으로 관리하며, 사용자가 프리셋 단위로 저장하고 적용할 수 있도록 다음과 같은 자료구조를 구성했습니다.
  
  <div align="center">
    <img width = "600" alt = "data Structure"  src = "https://github.com/user-attachments/assets/e499588b-e8b0-492b-9ea4-29f9f4b6ba91">
  </div>
</details>

<details>
  <summary>
    <strong>2-2. 비동기 작업의 관리</strong>
  </summary>
  
  `async/ await` 구문을 효과적으로 사용해야 했으며, 예상치 못한 사태를 방지하기 위해 `try/ catch` 구문을 적극적으로 활용했습니다.
  
  예를 들어, 유효성 탭의 수식 테스트 기능은 대표적으로 비동기적 흐름을 조율하기 위해 주의를 기울인 기능이었으며 다음과 같은 순서로 진행됐습니다.
  
  <div align="center">
    <table>
      <tr>
        <td>
          <img width="400" alt = "async/await code" src="https://github.com/user-attachments/assets/ca856087-058a-4a7f-8c47-b27c7bbce517">
        </td>
        <td>
          <img width="400" alt = "async/await image"src="https://github.com/user-attachments/assets/a1ae113d-f81a-48f0-848d-b0ba83f26ae5">
        </td>
      </td>
    </table>
  </div>
  
  위와 같이 작업의 흐름을 명확히 하여 다음과 같은 효과를 의도하였습니다.
  
  - **동기화 시점 최적화**: 필요한 시점에만 `context.sync()`를 호출하여 불필요한 동기화 작업을 최소화했습니다.
  - **의도치 않은 에러 처리**: `try/catch` 구문의 사용으로 의도치 않은 에러를 파악하고, 사용자에게 Message Box를 팝업하여 Error 사항을 인지할 수 있도록 적절히 처리했습니다.
</details>

<br/>

## 🚨 기획 변경

### 1. 매크로 녹화 기능 구현 이슈

#### 1-1. 기존안 및 변경 원인

> [!NOTE]
> Excel 정식 매크로 녹화 기능: 사용자의 조작 내용을 **VBA(Visual Basic Application)** 코드로, 자동으로 변환하여 기록하는 기능입니다.

기존안은 Add-in과 `VBA`로 Excel Application의 매크로 녹화 기능을 작동시켜 Excel 정식 매크로 기능을 바탕으로 트리거 역할을 하는 기능을 만드는 것이었습니다.
초기에 의도한 대로 매크로 녹화 기능을 작동시키는 데 성공하였지만, 사용자가 인터페이스를 직접 클릭하지 않는 이상 `VBA` 코드 녹화가 진행되지 않는다는 문제점에 봉착했습니다.

#### 1-2. 발생 원인

VBA로 매크로 녹화를 시작하는 것은 프로그래밍적으로 녹화 기능을 켜는 것과 사용자가 직접 인터페이스상으로 "매크로 녹화" 버튼을 클릭하는 것을 다르게 취급합니다. Excel은 보안상의 이유로 프로그래밍 방식의 입력 모니터링을 제한하는 것을 발견했습니다.

#### 1-3. 변경안
  
정식 매크로 기능을 사용할 수 없으므로, `Office JS`만을 활용하여 매크로 기능을 구현했습니다.<br/>
이를 위해 다음과 같은 이벤트에 감지 및 기록 함수를 등록하였습니다.

- **시트 내용 변경 이벤트**: 셀 내용 변경을 기록합니다.
- **서식 변경 이벤트**: 셀 서식 변경을 기록합니다.
- **표 추가 이벤트**: 생성된 표의 데이터 원본을 기록합니다.
- **표 내용 변경 이벤트**: 변경된 표의 데이터 원본과 변경 내역을 기록합니다.
- **차트 추가 이벤트**: 생성된 차트의 데이터 원본과 차트 타입을 기록합니다.
  
  API가 지원하는 변경 감지 이벤트가 한정적이기 때문에 기존 정식 매크로 기능을 완벽히 구현할 순 없었지만, Excel 사용자가 주로 사용하는 변경 내역을 감지하고 순서대로 기록하여, 사용자가 조작한 내역을 재생할 수 있도록 구현하여 해결할 수 있었습니다.
</details>

<br/>

### 2. 추출하기 기능 구현 이슈

#### 2-1. 기존안 및 변경 원인

기존안은 `Office JS`의 `Outlook API`를 동시에 이용하여, 사용자가 희망하는 범위의 내용을 첨부파일로서 메일로 전송하는 편의 기능이었습니다.
하지만, Excel 내에서 Outlook을 실행시키거나 통신하는 것이 불가능하다는 문제를 깨달아 기획을 변경하게 됐습니다.

#### 2-2. 발생 원인

Add-in은 클라이언트 측에서 실행되는 JavaScript로 구동되기 때문에, 무분별한 접근이 가능할 경우 발생할 보안 문제를 예방하기 위해 지원하는 기능이 제한되어 있습니다.
또한, Office JS API는 여전히 발전 중인 새로운 기술로, 현재 점진적으로 API를 확장해 나가고 있는 상황입니다.

#### 2-3. 변경안
  
  1. 초기 기획 의도 유지<br/>
     기획 의도가 메인 타겟인 사무직 분들의 특정 부분을 메일로 요청받을 경우의 불편을 해소하기 위함이었던 만큼, 해당 불편함을 다른 방도로 해소하기 위해 사용자가 희망하는 범위를 새로운 Excel 통합 문서로 추출하는 기능으로 변경했습니다.
  
     하지만 `Office JS`로 구현하기엔 다음과 같은 제약 사항이 있었습니다.
  
     - 모든 속성을 전부 로드, 저장해야 함.
     - 완벽히 복사, 붙여넣기엔 한정된 API 지원 속성.
  
     이런 한계점을 극복하기 위해 `VBA`를 활용했습니다.
  
> [!NOTE]
> **VBA(Visual Basic for Application)**: MS Office에 내장된 프로그래밍 언어로, Excel에 깊숙이 통합되어 있어 다양한 기능들을 더욱 자유롭게 구현할 수 있습니다.
>
> ### 하지만 직접적인 `Add-in`과 `VBA` 간의 상호작용은 지원되지 않습니다!
  
  2. Add-in과 `VBA`의 연계<br/>
     정식적으로 호환되지 않는 두 기능이므로 다음과 같이 우회하는 방법을 사용했습니다.
  
  <div align="center">
    <img width="300" alt="excel with VBA" src="https://github.com/user-attachments/assets/5386073b-2435-410f-9deb-e7d28c347c9b">
  </div>
  
  또한, `VBA` 함수는 사용자가 직접 등록해야 하므로 사용자가 추출하기 기능을 사용하기 전에, 사용 방법 및 해당 `VBA` 추가 기능 등록 파일을 담은 Notion 사이트를 이용할 수 있도록 Message Box를 볼 수 있도록 했습니다.<br/>
  
  <div align="center">
    <table>
      <tr>
        <td>
          <img width="400" alt = "alert message box" src="https://github.com/user-attachments/assets/a738f6fd-aa42-4e4d-b579-e0d385a43f16">
        </td>
        <td>
          <img width="400" alt="how install vba" src="https://github.com/user-attachments/assets/5519eb64-8953-4b57-bc59-8d9a825ba26f"><br/>
          <a href="https://github.com/user-attachments/assets/5519eb64-8953-4b57-bc59-8d9a825ba26f">VBA 추가 안내서</a>
        </td>
      </tr>
    </table>
  </div>
  
  이처럼 두 가지 프로그래밍 언어를 병용하여 사용하는 방식으로 해당 문제를 해결할 수 있었습니다.
</details>

<br/>

## ⌛️ 회고

모든 과정을 혼자서 소화해야 하는 개인 프로젝트는 생각했던 것 이상으로 촉박한 일정과 강행군이었습니다.
Office JS를 주축으로 진행하는 프로젝트인 이상 사용법 파악과 문제 해결은 전부 제 개인의 몫이었습니다.<br/><br/>
프로젝트를 진행하며 컨셉이나 구현 방향성에 관해 토론하고 의견을 나누는 과정이 없었기에 막상 구현할 때가 되어서 처음 기획 의도와 달라 당황하게 된 경험을 했습니다.<br/>
이번 프로젝트를 바탕으로 기획 단계에서 좀 더 계획 및 구현 방안을 구체화하는 것의 중요성을 깨달았습니다.<br/><br/>
 하지만 동시에, 코딩의 코 자도 모르던 3개월 전과는 다르게 저만의 힘으로 이와 같은 프로젝트를 완성 시켰다는 뿌듯함 또한 느낄 수 있었습니다.<br/>
3개월로 이뤄낸 성과가 무색해지지 않도록 노력해 나갈 수 있는 원동력을 얻게 된 시간이었습니다.
