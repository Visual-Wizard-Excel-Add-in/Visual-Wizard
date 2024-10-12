const MESSAGE_LIST = {
  loadFail: {
    type: "warning",
    title: "Load Failed:",
    body: "데이터를 불러오는데 실패했습니다.",
  },
  saveFail: {
    type: "warning",
    title: "Save Failed:",
    body: "데이터를 저장하는데 실패했습니다.",
  },
  workFail: {
    type: "warning",
    title: "Work Failed",
    body: "실행에 실패했습니다.",
  },
  saveSuccess: {
    type: "success",
    title: "Saved",
    body: "데이터를 저장했습니다.",
  },
  loadSuccess: {
    type: "success",
    title: "Loaded",
    body: "데이터를 불러왔습니다.",
  },
  default: {
    type: "warning",
    title: "Undefiend Error:",
    body: "예상하지 못한 에러가 발생했습니다.",
  },
};

export default MESSAGE_LIST;
