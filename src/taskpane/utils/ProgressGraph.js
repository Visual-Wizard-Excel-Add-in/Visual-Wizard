import ProgressGraphNode from "./ProgressGraphNode";

class ProgressGraph {
  constructor() {
    this.nodes = new Map();
  }

  addNode(node) {
    const key = node.formula || node.address;

    if (!this.nodes.has(key)) {
      this.nodes.set(key, new ProgressGraphNode(node));
    }

    return this.nodes.get(key);
  }

  addDependency(from, to) {
    const fromNode = this.addNode(from);
    const toNode = this.addNode(to);

    toNode.dependencies.add(fromNode);
  }

  topologicalSort() {
    const sorted = [];
    const visited = new Set();

    const visit = (node) => {
      if (visited.has(node.data.formula)) return;

      visited.add(node.data.formula);

      Array.from(node.dependencies).forEach((dep) => visit(dep));

      sorted.unshift(node.data);
    };

    this.nodes.forEach((node) => {
      if (!visited.has(node.data.formula)) {
        visit(node);
      }
    });

    return sorted;
  }
}

export default ProgressGraph;