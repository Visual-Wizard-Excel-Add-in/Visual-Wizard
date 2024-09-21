class ProgressGraphNode {
  data: GraphNodeType;
  dependencies: Set<GraphType>;

  constructor(data: GraphNodeType) {
    this.data = data;
    this.dependencies = new Set();
  }
}

export default ProgressGraphNode;
