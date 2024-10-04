class Message {
  constructor(purpose, option = null) {
    this.purpose = purpose;
    this.option = option;
  }

  get body() {
    switch (this.purpose) {
      case "loadFail":
        return {
          type: "warning",
          title: "Load Failed:",
          body: `데이터를 불러오는데 실패했습니다.\n${this.option}`,
        };

      case "saveFail":
        return {
          type: "warning",
          title: "Save Failed:",
          body: `데이터를 저장하는데 실패했습니다.\n${this.option}`,
        };

      case "workFail":
        return {
          type: "warning",
          title: "Work Failed",
          body: `실행에 실패했습니다.\n${this.option}`,
        };

      case "saveSuccess":
        return {
          type: "success",
          title: "Saved",
          body: `데이터를 저장했습니다.\n${this.option}`,
        };

      case "loadSuccess":
        return {
          type: "success",
          title: "Loaded",
          body: `데이터를 불러왔습니다.\n${this.option}`,
        };

      default:
        return {
          type: "warning",
          title: "Undefiend Error:",
          body: `예상하지 못한 에러가 발생했습니다.\n${this.option}`,
        };
    }
  }
}

export default Message;
