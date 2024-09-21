declare const OfficeRuntime: typeof import("office-runtime");

interface GraphNodeType {
  address: string;
  condition?: string;
  dependencies?: GraphNodeType[];
  falseValue?: string;
  formula: string;
  functionName?: string;
  trueValue?: string;
  values?: string[];
  criteriaRange?: string;
  criteria?: string | string[];
  sumRange?: string;
  conditions?: string[];
  criteriaRanges?: string[];
}

interface GraphType {
  data: GraphNodeType;
  dependencies: Set<GraphType>;
}
