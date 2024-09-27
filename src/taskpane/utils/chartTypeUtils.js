const CHART_TYPE_LIST = [
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

function translateChartTypeKOR(chartTypeEnglish) {
  switch (chartTypeEnglish) {
    case "Invalid":
      return "유효하지 않음";

    case "ColumnClustered":
      return "묶은 세로 막대형 차트";

    case "ColumnStacked":
      return "누적 세로 막대형 차트";

    case "ColumnStacked100":
      return "100% 기준 누적 세로 막대형 차트";

    case "3DColumnClustered":
      return "3D 묶은 세로 막대형 차트";

    case "3DColumnStacked":
      return "3D 누적 세로 막대형 차트";

    case "3DColumnStacked100":
      return "3D 100% 기준 누적 세로 막대형 차트";

    case "3DColumn":
      return "3차원 세로 막대형 차트";

    case "BarClustered":
      return "묶은 가로 막대형 차트";

    case "BarStacked":
      return "누적 가로 막대형 차트";

    case "BarStacked100":
      return "100% 기준 누적 가로 막대형 차트";

    case "3DBarClustered":
      return "3차원 묶은 가로 막대형 차트";

    case "3DBarStacked":
      return "3차원 누적 가로 막대형 차트";

    case "3DBarStacked100":
      return "3차원 100% 기준 누적 가로 막대형 차트";

    case "Line":
      return "꺾은선형 차트";

    case "LineStacked":
      return "누적 꺾은선형 차트";

    case "LineStacked100":
      return "100% 기준 누적 꺾은선형 차트";

    case "LineMarkers":
      return "표식이 있는 꺾은선형 차트";

    case "LineMarkersStacked":
      return "표식이 있는 누적 꺾은선형 차트";

    case "LineMarkersStacked100":
      return "표식이 있는 100% 기준 누적 꺾은선형 차트";

    case "3DLine":
      return "3D 꺾은선형 차트";

    case "Area":
      return "영역형 차트";

    case "AreaStacked":
      return "누적 영역형 차트";

    case "AreaStacked100":
      return "100% 기준 누적 영역형 차트";

    case "3DArea":
      return "3D 영역형 차트";

    case "3DAreaStacked":
      return "3D 누적 영역형 차트";

    case "3DAreaStacked100":
      return "3D 100% 기준 누적 영역형 차트";

    case "Pie":
      return "원형 차트";

    case "PieOfPie":
      return "원형 대 원형 차트";

    case "BarOfPie":
      return "원형 대 가로 막대형 차트";

    case "3DPie":
      return "3차원 원형 차트";

    case "PieExploded":
      return "분리형 원형 차트";

    case "3DPieExploded":
      return "3D 분리형 원형 차트";

    case "Doughnut":
      return "도넛형 차트";

    case "DoughnutExploded":
      return "분리형 도넛형 차트";

    case "Treemap":
      return "트리맵 차트";

    case "Sunburst":
      return "선버스트 차트";

    case "Histogram":
      return "히스토그램형 차트";

    case "Pareto":
      return "파레토 차트";

    case "BoxWhisker":
      return "상자 수염 차트";

    case "XYScatter":
      return "분산형 차트";

    case "XYScatterSmooth":
      return "곡선 및 표식이 있는 분산형 차트";

    case "XYScatterSmoothNoMarkers":
      return "곡선 및 표식이 없는 분산형 차트";

    case "XYScatterLines":
      return "직선 및 표식이 있는 분산형 차트";

    case "XYScatterLinesNoMarkers":
      return "직선 및 표식이 없는 분산형 차트";

    case "Bubble":
      return "거품형 차트";

    case "Bubble3DEffect":
      return "3차원 거품형 차트";

    case "Waterfall":
      return "폭포 차트";

    case "Funnel":
      return "깔때기형 차트";

    case "StockHLC":
      return "고가-저가-종가 차트";

    case "StockOHLC":
      return "시가-고가-저가-종가 차트";

    case "StockVHLC":
      return "거래량-고가-저가-종가 차트";

    case "StockVOHLC":
      return "거래량-시가-고가-저가-종가 차트";

    case "Surface":
      return "3차원 표면형 차트";

    case "SurfaceWireframe":
      return "3차원 표면형 차트(골격형)";

    case "SurfaceTopView":
      return "표면형 차트(조감도)";

    case "SurfaceTopViewWireframe":
      return "표면형 차트(골격형 조감도)";

    case "Radar":
      return "방사형 차트";

    case "RadarMarkers":
      return "표식이 있는 방사형 차트";

    case "RadarFilled":
      return "채워진 방사형 차트";

    case "CylinderColClustered":
      return "원기둥 묶은 세로 막대형 차트";

    case "CylinderColStacked":
      return "원기둥 누적 세로 막대형 차트";

    case "CylinderColStacked100":
      return "원기둥 100% 기준 누적 세로 막대형 차트";

    case "CylinderBarClustered":
      return "원기둥 묶은 가로 막대형 차트";

    case "CylinderBarStacked":
      return "원기둥 누적 가로 막대형 차트";

    case "CylinderBarStacked100":
      return "원기둥 100% 기준 누적 가로 막대형 차트";

    case "CylinderCol":
      return "원기둥 세로 막대형 차트";

    case "ConeColClustered":
      return "원뿔 묶은 세로 막대형 차트";

    case "ConeColStacked":
      return "원뿔 누적 세로 막대형 차트";

    case "ConeColStacked100":
      return "원뿔 100% 기준 누적 세로 막대형 차트";

    case "ConeBarClustered":
      return "원뿔 묶은 가로 막대형 차트";

    case "ConeBarStacked":
      return "원뿔 누적 가로 막대형 차트";

    case "ConeBarStacked100":
      return "원뿔 100% 기준 누적 가로 막대형 차트";

    case "ConeCol":
      return "원뿔 세로 막대형 차트";

    case "PyramidColClustered":
      return "피라미드 묶은 세로 막대형 차트";

    case "PyramidColStacked":
      return "피라미드 누적 세로 막대형 차트";

    case "PyramidColStacked100":
      return "피라미드 100% 기준 누적 세로 막대형 차트";

    case "PyramidBarClustered":
      return "피라미드 묶은 가로 막대형 차트";

    case "PyramidBarStacked":
      return "피라미드 누적 가로 막대형 차트";

    case "PyramidBarStacked100":
      return "피라미드 100% 기준 누적 가로 막대형 차트";

    case "PyramidCol":
      return "피라미드 세로 막대형 차트";

    case "RegionMap":
      return "지역 맵형 차트";

    default:
      return "알 수 없는 차트 유형";
  }
}

function translateChartTypeENG(chartTypeKorean) {
  switch (chartTypeKorean) {
    case "유효하지 않음":
      return "Invalid";

    case "묶은 세로 막대형 차트":
      return "ColumnClustered";

    case "누적 세로 막대형 차트":
      return "ColumnStacked";

    case "100% 기준 누적 세로 막대형 차트":
      return "ColumnStacked100";

    case "3D 묶은 세로 막대형 차트":
      return "3DColumnClustered";

    case "3D 누적 세로 막대형 차트":
      return "3DColumnStacked";

    case "3D 100% 기준 누적 세로 막대형 차트":
      return "3DColumnStacked100";

    case "3차원 세로 막대형 차트":
      return "3DColumn";

    case "묶은 가로 막대형 차트":
      return "BarClustered";

    case "누적 가로 막대형 차트":
      return "BarStacked";

    case "100% 기준 누적 가로 막대형 차트":
      return "BarStacked100";

    case "3차원 묶은 가로 막대형 차트":
      return "3DBarClustered";

    case "3차원 누적 가로 막대형 차트":
      return "3DBarStacked";

    case "3차원 100% 기준 누적 가로 막대형 차트":
      return "3DBarStacked100";

    case "꺾은선형 차트":
      return "Line";

    case "누적 꺾은선형 차트":
      return "LineStacked";

    case "100% 기준 누적 꺾은선형 차트":
      return "LineStacked100";

    case "표식이 있는 꺾은선형 차트":
      return "LineMarkers";

    case "누적 꺾은선형 표식 차트":
      return "LineMarkersStacked";

    case "100% 기준 누적 꺾은선형 표식 차트":
      return "LineMarkersStacked100";

    case "3D 꺾은선형 차트":
      return "3DLine";

    case "영역형 차트":
      return "Area";

    case "누적 영역형 차트":
      return "AreaStacked";

    case "100% 기준 누적 영역형 차트":
      return "AreaStacked100";

    case "3D 영역형 차트":
      return "3DArea";

    case "3D 누적 영역형 차트":
      return "3DAreaStacked";

    case "3D 100% 기준 누적 영역형 차트":
      return "3DAreaStacked100";

    case "원형 차트":
      return "Pie";

    case "원형 대 원형 차트":
      return "PieOfPie";

    case "원형 대 가로 막대형 차트":
      return "BarOfPie";

    case "3차원 원형 차트":
      return "3DPie";

    case "분리형 원형 차트":
      return "PieExploded";

    case "3D 분리형 원형 차트":
      return "3DPieExploded";

    case "도넛형 차트":
      return "Doughnut";

    case "분리형 도넛형 차트":
      return "DoughnutExploded";

    case "트리맵 차트":
      return "Treemap";

    case "선버스트 차트":
      return "Sunburst";

    case "히스토그램형 차트":
      return "Histogram";

    case "파레토 차트":
      return "Pareto";

    case "상자 수염 차트":
      return "BoxWhisker";

    case "분산형 차트":
      return "XYScatter";

    case "곡선 및 표식이 있는 분산형 차트":
      return "XYScatterSmooth";

    case "곡선 및 표식이 없는 분산형 차트":
      return "XYScatterSmoothNoMarkers";

    case "직선 및 표식이 있는 분산형 차트":
      return "XYScatterLines";

    case "직선 및 표식이 없는 분산형 차트":
      return "XYScatterLinesNoMarkers";

    case "거품형 차트":
      return "Bubble";

    case "3차원 거품형 차트":
      return "Bubble3DEffect";

    case "폭포 차트":
      return "Waterfall";

    case "깔때기형 차트":
      return "Funnel";

    case "고가-저가-종가 차트":
      return "StockHLC";

    case "시가-고가-저가-종가 차트":
      return "StockOHLC";

    case "거래량-고가-저가-종가 차트":
      return "StockVHLC";

    case "거래량-시가-고가-저가-종가 차트":
      return "StockVOHLC";

    case "3차원 표면형 차트":
      return "Surface";

    case "3차원 표면형 차트(골격형)":
      return "SurfaceWireframe";

    case "표면형 차트(조감도)":
      return "SurfaceTopView";

    case "표면형 차트(골격형 조감도)":
      return "SurfaceTopViewWireframe";

    case "방사형 차트":
      return "Radar";

    case "표식이 있는 방사형 차트":
      return "RadarMarkers";

    case "채워진 방사형 차트":
      return "RadarFilled";

    case "원기둥 묶은 세로 막대형 차트":
      return "CylinderColClustered";

    case "원기둥 누적 세로 막대형 차트":
      return "CylinderColStacked";

    case "원기둥 100% 기준 누적 세로 막대형 차트":
      return "CylinderColStacked100";

    case "원기둥 묶은 가로 막대형 차트":
      return "CylinderBarClustered";

    case "원기둥 누적 가로 막대형 차트":
      return "CylinderBarStacked";

    case "원기둥 100% 기준 누적 가로 막대형 차트":
      return "CylinderBarStacked100";

    case "원기둥 세로 막대형 차트":
      return "CylinderCol";

    case "원뿔 묶은 세로 막대형 차트":
      return "ConeColClustered";

    case "원뿔 누적 세로 막대형 차트":
      return "ConeColStacked";

    case "원뿔 100% 기준 누적 세로 막대형 차트":
      return "ConeColStacked100";

    case "원뿔 묶은 가로 막대형 차트":
      return "ConeBarClustered";

    case "원뿔 누적 가로 막대형 차트":
      return "ConeBarStacked";

    case "원뿔 100% 기준 누적 가로 막대형 차트":
      return "ConeBarStacked100";

    case "원뿔 세로 막대형 차트":
      return "ConeCol";

    case "피라미드 묶은 세로 막대형 차트":
      return "PyramidColClustered";

    case "피라미드 누적 세로 막대형 차트":
      return "PyramidColStacked";

    case "피라미드 100% 기준 누적 세로 막대형 차트":
      return "PyramidColStacked100";

    case "피라미드 묶은 가로 막대형 차트":
      return "PyramidBarClustered";

    case "피라미드 누적 가로 막대형 차트":
      return "PyramidBarStacked";

    case "피라미드 100% 기준 누적 가로 막대형 차트":
      return "PyramidBarStacked100";

    case "피라미드 세로 막대형 차트":
      return "PyramidCol";

    case "지역 맵형 차트":
      return "RegionMap";

    default:
      return "Unknown chart type";
  }
}

export { CHART_TYPE_LIST, translateChartTypeKOR, translateChartTypeENG };
