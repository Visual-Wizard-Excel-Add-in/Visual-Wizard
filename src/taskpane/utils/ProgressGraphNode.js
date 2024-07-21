class ProgressGraphNode {
  constructor(data) {
    this.data = data;
    this.dependencies = new Set();
  }
}

export default ProgressGraphNode;
