const chartTypeList = [
  { value: "Invalid", name: "유효하지 않음", label: "기타" },
  {
    value: "ColumnClustered",
    name: "묶은 세로 막대형 차트",
    label: "2차원 세로 막대형",
  },
  {
    value: "ColumnStacked",
    name: "누적 세로 막대형 차트",
    label: "2차원 세로 막대형",
  },
  {
    value: "ColumnStacked100",
    name: "100% 기준 누적 세로 막대형 차트",
    label: "2차원 세로 막대형",
  },
  {
    value: "3DColumnClustered",
    name: "3D 묶은 세로 막대형 차트",
    label: "3차원 세로 막대형",
  },
  {
    value: "3DColumnStacked",
    name: "3D 누적 세로 막대형 차트",
    label: "3차원 세로 막대형",
  },
  {
    value: "3DColumnStacked100",
    name: "3D 100% 기준 누적 세로 막대형 차트",
    label: "3차원 세로 막대형",
  },
  {
    value: "3DColumn",
    name: "3차원 세로 막대형 차트",
    label: "3차원 세로 막대형",
  },

  {
    value: "BarClustered",
    name: "묶은 가로 막대형 차트",
    label: "2차원 가로 막대형",
  },
  {
    value: "BarStacked",
    name: "누적 가로 막대형 차트",
    label: "2차원 가로 막대형",
  },
  {
    value: "BarStacked100",
    name: "100% 기준 누적 가로 막대형 차트",
    label: "2차원 가로 막대형",
  },
  {
    value: "3DBarClustered",
    name: "3차원 묶은 가로 막대형 차트",
    label: "3차원 가로 막대형",
  },
  {
    value: "3DBarStacked",
    name: "3차원 누적 가로 막대형 차트",
    label: "3차원가로 막대형",
  },
  {
    value: "3DBarStacked100",
    name: "3차원 100% 기준 누적 가로 막대형 차트",
    label: "3차원 가로 막대형",
  },
  { value: "Line", name: "꺾은선형 차트", label: "2차원 꺾은선형" },
  { value: "LineStacked", name: "누적 꺾은선형 차트", label: "2차원 꺾은선형" },
  {
    value: "LineStacked100",
    name: "100% 기준 누적 꺾은선형 차트",
    label: "2차원 꺾은선형",
  },
  {
    value: "LineMarkers",
    name: "표식이 있는 꺾은선형 차트",
    label: "2차원 꺾은선형",
  },
  {
    value: "LineMarkersStacked",
    name: "표식이 있는 누적 꺾은선형 차트",
    label: "2차원 꺾은선형",
  },
  {
    value: "LineMarkersStacked100",
    name: "표식이 있는 100% 기준 누적 꺾은선형 차트",
    label: "2차원 꺾은선형",
  },
  { value: "3DLine", name: "3D 꺾은선형 차트", label: "3차원 꺾은선형" },
  { value: "Area", name: "영역형 차트", label: "2차원 영역형" },
  { value: "AreaStacked", name: "누적 영역형 차트", label: "2차원 영역형" },
  {
    value: "AreaStacked100",
    name: "100% 기준 누적 영역형 차트",
    label: "2차원 영역형",
  },
  { value: "3DArea", name: "3D 영역형 차트", label: "3차원 영역형" },
  {
    value: "3DAreaStacked",
    name: "3D 누적 영역형 차트",
    label: "3차원 영역형",
  },
  {
    value: "3DAreaStacked100",
    name: "3D 100% 기준 누적 영역형 차트",
    label: "3차원 영역형",
  },
  { value: "Pie", name: "원형 차트", label: "2차원 원형" },
  { value: "PieOfPie", name: "원형 대 원형 차트", label: "2차원 원형" },
  { value: "BarOfPie", name: "원형 대 가로 막대형 차트", label: "2차원 원형" },
  { value: "3DPie", name: "3차원 원형 차트", label: "3차원 원형" },
  { value: "PieExploded", name: "분리형 원형 차트", label: "원형" },
  { value: "3DPieExploded", name: "3D 분리형 원형 차트", label: "원형" },
  { value: "Doughnut", name: "도넛형 차트", label: "도넛형" },
  { value: "Treemap", name: "트리맵 차트", label: "트리맵" },
  { value: "Sunburst", name: "선버스트 차트", label: "선버스트" },
  { value: "Histogram", name: "히스토그램형 차트", label: "히스토그램" },
  { value: "Pareto", name: "파레토 차트", label: "히스토그램" },
  { value: "Boxwhisker", name: "상자 수염 차트", label: "상자 수염" },
  { value: "XYScatter", name: "분산형 차트", label: "분산형" },
  {
    value: "XYScatterSmooth",
    name: "곡선 및 표식이 있는 분산형 차트",
    label: "분산형",
  },
  {
    value: "XYScatterSmoothNoMarkers",
    name: "곡선 및 표식이 없는 분산형 차트",
    label: "분산형",
  },
  {
    value: "XYScatterLines",
    name: "직선 및 표식이 있는 분산형 차트",
    label: "분산형",
  },
  {
    value: "XYScatterLinesNoMarkers",
    name: "직선 및 표식이 없는 분산형 차트",
    label: "분산형",
  },
  { value: "Bubble", name: "거품형 차트", label: "거품형" },
  { value: "Bubble3DEffect", name: "3차원 거품형 차트", label: "거품형" },
  { value: "Waterfall", name: "폭포 차트", label: "폭포" },
  { value: "Funnel", name: "깔때기형 차트", label: "깔때기형" },
  { value: "StockHLC", name: "고가-저가-종가 차트", label: "주식형" },
  { value: "StockOHLC", name: "시가-고가-저가-종가 차트", label: "주식형" },
  { value: "StockVHLC", name: "거래량-고가-저가-종가 차트", label: "주식형" },
  {
    value: "StockVOHLC",
    name: "거래량-시가-고가-저가-종가 차트",
    label: "주식형",
  },
  { value: "Surface", name: "3차원 표면형 차트", label: "표면형" },
  {
    value: "SurfaceWireframe",
    name: "3차원 표면형 차트(골격형)",
    label: "표면형",
  },
  { value: "SurfaceTopView", name: "표면형 차트(조감도)", label: "표면형" },
  {
    value: "SurfaceTopViewWireframe",
    name: "표면형 차트(골격형 조감도)",
    label: "표면형",
  },
  { value: "Radar", name: "방사형 차트", label: "방사형" },
  {
    value: "RadarMarkers",
    name: "표식이 있는 방사형 차트",
    label: "방사형",
  },
  { value: "RadarFilled", name: "채워진 방사형 차트", label: "방사형" },

  { value: "DoughnutExploded", name: "분리형 도넛형 차트", label: "도넛형" },

  {
    value: "CylinderColClustered",
    name: "원기둥 묶은 세로 막대형 차트",
    label: "원기둥형",
  },
  {
    value: "CylinderColStacked",
    name: "원기둥 누적 세로 막대형 차트",
    label: "원기둥형",
  },
  {
    value: "CylinderColStacked100",
    name: "원기둥 100% 기준 누적 세로 막대형 차트",
    label: "원기둥형",
  },
  {
    value: "CylinderBarClustered",
    name: "원기둥 묶은 가로 막대형 차트",
    label: "원기둥형",
  },
  {
    value: "CylinderBarStacked",
    name: "원기둥 누적 가로 막대형 차트",
    label: "원기둥형",
  },
  {
    value: "CylinderBarStacked100",
    name: "원기둥 100% 기준 누적 가로 막대형 차트",
    label: "원기둥형",
  },
  { value: "CylinderCol", name: "원기둥 세로 막대형 차트", label: "원기둥형" },
  {
    value: "ConeColClustered",
    name: "원뿔 묶은 세로 막대형 차트",
    label: "원뿔형",
  },
  {
    value: "ConeColStacked",
    name: "원뿔 누적 세로 막대형 차트",
    label: "원뿔형",
  },
  {
    value: "ConeColStacked100",
    name: "원뿔 100% 기준 누적 세로 막대형 차트",
    label: "원뿔형",
  },
  {
    value: "ConeBarClustered",
    name: "원뿔 묶은 가로 막대형 차트",
    label: "원뿔형",
  },
  {
    value: "ConeBarStacked",
    name: "원뿔 누적 가로 막대형 차트",
    label: "원뿔형",
  },
  {
    value: "ConeBarStacked100",
    name: "원뿔 100% 기준 누적 가로 막대형 차트",
    label: "원뿔형",
  },
  { value: "ConeCol", name: "원뿔 세로 막대형 차트", label: "원뿔형" },
  {
    value: "PyramidColClustered",
    name: "피라미드 묶은 세로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidColStacked",
    name: "피라미드 누적 세로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidColStacked100",
    name: "피라미드 100% 기준 누적 세로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidBarClustered",
    name: "피라미드 묶은 가로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidBarStacked",
    name: "피라미드 누적 가로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidBarStacked100",
    name: "피라미드 100% 기준 누적 가로 막대형 차트",
    label: "피라미드형",
  },
  {
    value: "PyramidCol",
    name: "피라미드 세로 막대형 차트",
    label: "피라미드형",
  },

  { value: "RegionMap", name: "지역 맵형 차트", label: "맵형" },
];

export default chartTypeList;
