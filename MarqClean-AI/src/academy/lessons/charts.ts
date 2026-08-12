import type { Lesson } from "./types";

export const charts: Lesson[] = [
  {
    slug: "create-first-chart",
    title: "Create your first chart and choose the right type",
    track: "charts",
    level: "Beginner",
    summary:
      "One keystroke makes a chart. The skill is choosing a type that answers the question your reader is actually asking.",
    objectives: [
      "Insert a chart from a range in one keystroke",
      "Pick the correct chart type for the message",
      "Use the three chart buttons and Recommended Charts",
    ],
    steps: [
      {
        title: "Lay the data out simply",
        detail:
          "Labels in the first column, values in the columns to the right, one header row, no blank rows or merged cells.",
      },
      {
        title: "Insert it",
        detail:
          "Select the data including headers and press Alt + F1 for a chart on the same sheet, or F11 for a chart on its own sheet. Insert > Recommended Charts suggests sensible options.",
      },
      {
        title: "Choose by the question",
        detail:
          "Comparison between items: column or bar. Change over time: line. Part of a whole: pie or 100% stacked bar. Relationship between two measures: scatter. Distribution: histogram. Progress towards a target: combo with a line.",
      },
      {
        title: "Bar instead of column",
        detail:
          "When category names are long, or there are more than about eight of them, a horizontal bar chart is easier to read.",
      },
      {
        title: "The three buttons",
        detail:
          "Select the chart and use the icons on its right: plus adds elements (titles, labels, legend), brush changes style and colour, funnel filters which series and categories are plotted.",
      },
      {
        title: "Switch rows and columns",
        detail:
          "Chart Design > Switch Row/Column flips what is compared with what. Try it before you rebuild anything.",
      },
      {
        title: "Move and resize",
        detail:
          "Chart Design > Move Chart puts it on its own sheet. Hold Alt while dragging to snap the chart to the cell gridlines.",
      },
      {
        title: "Make it live",
        detail:
          "Build the chart on a Table (Ctrl + T) and it grows automatically as new rows arrive.",
      },
    ],
    example: {
      caption: "Which chart answers which question",
      headers: ["Question", "Chart type"],
      rows: [
        ["Who sold the most?", "Column or bar"],
        ["How has revenue moved this year?", "Line"],
        ["What share does each region hold?", "Pie or 100% stacked bar"],
        ["Does spend drive sales?", "Scatter with trendline"],
        ["How are order values distributed?", "Histogram"],
        ["Actual against target", "Combo: columns plus a line"],
      ],
    },
    shortcuts: [
      ["Alt + F1", "Chart on the current sheet"],
      ["F11", "Chart on a new sheet"],
      ["Ctrl + 1", "Format the selected chart element"],
    ],
    practice: {
      task: "Twelve months of revenue need a chart that shows the trend and highlights the best month.",
      answer:
        "Select the data, Alt + F1, change to a Line chart, then click the best point twice to select it alone and add a data label.",
    },
    quiz: [
      {
        q: "Which chart type shows change over time best?",
        options: ["Pie", "Line", "Doughnut", "Radar"],
        answer: 1,
        explain: "A line chart makes trends and turning points obvious.",
      },
    ],
  },
  {
    slug: "format-charts",
    title: "Formatting charts so they communicate",
    track: "charts",
    level: "Intermediate",
    summary:
      "Axis limits, data labels, gridlines and colour. Small formatting choices decide whether a chart informs or misleads.",
    objectives: [
      "Format every chart element with one dialog",
      "Control axis scale, units and number format",
      "Strip the clutter and highlight the point",
    ],
    steps: [
      {
        title: "Select then Ctrl + 1",
        detail:
          "Click any element (axis, series, title, gridline) and press Ctrl + 1 to open its format pane. The pane changes to match whatever is selected.",
      },
      {
        title: "Axis scale honesty",
        detail:
          "Format Axis > Bounds sets minimum and maximum. A column chart should start at zero; truncating the axis exaggerates differences. A line chart may start elsewhere if you label it clearly.",
      },
      {
        title: "Display units",
        detail:
          "Format Axis > Display units > Thousands or Millions shortens crowded labels, and you can show the unit caption automatically.",
      },
      {
        title: "Number format on the axis",
        detail:
          "The same custom codes as cells work here, for example #,##0\" kg\" or 0.0%.",
      },
      {
        title: "Data labels",
        detail:
          "Plus button > Data Labels. In the format pane tick 'Value From Cells' to label points with text from any range, and use the Number section to format them.",
      },
      {
        title: "Remove the clutter",
        detail:
          "Delete the legend when there is only one series, thin or delete the gridlines, remove the chart border, and put the title in plain words that state the finding.",
      },
      {
        title: "Highlight one point",
        detail:
          "Click a series once to select all of it, then click again on a single column to select only that one, and give it a stronger colour. Everything else stays grey.",
      },
      {
        title: "Handle gaps",
        detail:
          "Chart Design > Select Data > Hidden and Empty Cells lets you choose Gaps, Zero or 'Connect data points with line' for missing values.",
      },
      {
        title: "Save the look",
        detail:
          "Right-click the chart > Save as Template, then apply it to future charts from Insert > Charts > Templates.",
      },
    ],
    practice: {
      task: "A column chart of monthly revenue is unreadable because the axis shows 0 to 14,000,000.",
      answer:
        "Format Axis > Display units > Millions, set the number format to 0.0\"m\", delete the gridlines and the legend, and write a title that states the finding.",
    },
    quiz: [
      {
        q: "Why should a column chart's value axis start at zero?",
        options: [
          "It looks tidier",
          "Otherwise the height differences exaggerate the real difference",
          "Excel requires it",
          "To fit the labels",
        ],
        answer: 1,
        explain: "Column length is the visual message, so a truncated axis misleads.",
      },
    ],
  },
  {
    slug: "advanced-charts",
    title: "Combo charts, secondary axes, trendlines and sparklines",
    track: "charts",
    level: "Advanced",
    summary:
      "Plot two very different measures together, add a forecast line, and squeeze a whole trend into a single cell.",
    objectives: [
      "Build a combo chart with a secondary axis",
      "Add and interpret a trendline",
      "Insert sparklines and in-cell bars",
    ],
    steps: [
      {
        title: "Combo chart",
        detail:
          "Select the data, Insert > Combo > Custom Combination. Choose Clustered Column for the amounts and Line for the percentage series.",
      },
      {
        title: "Secondary axis",
        detail:
          "In the same dialog tick 'Secondary Axis' next to the series with the different scale, such as a margin percentage against revenue in thousands.",
      },
      {
        title: "Use it sparingly",
        detail:
          "Two axes can imply a relationship that does not exist. Always label both axes and consider two stacked charts instead.",
      },
      {
        title: "Trendline",
        detail:
          "Right-click a series > Add Trendline. Linear for steady growth, Moving Average to smooth noise, and 'Forecast forward' to project ahead. Tick 'Display R-squared' to show how well it fits.",
      },
      {
        title: "Error bars",
        detail:
          "Plus button > Error Bars > More Options for fixed, percentage, standard deviation or custom values from a range.",
      },
      {
        title: "Sparklines",
        detail:
          "Insert > Sparklines > Line, choose the data range and the location cell. One tiny chart per row, perfect beside a KPI table. Use the Sparkline tab to mark High Point and Low Point.",
      },
      {
        title: "In-cell bars without a chart",
        detail:
          "Conditional Formatting > Data Bars, or the REPT trick for a text bar.",
        formula: '=REPT("|";ROUND(B2/MAX($B$2:$B$50)*20;0))',
      },
      {
        title: "Other useful types",
        detail:
          "Waterfall for bridging figures, Histogram for distribution, Box and Whisker for spread, Treemap and Sunburst for hierarchies, Map for country or region data, Funnel for pipelines.",
      },
    ],
    practice: {
      task: "Show monthly revenue in the thousands and gross margin percentage on the same chart.",
      answer:
        "Insert > Combo > Custom Combination: Revenue as Clustered Column, Margin % as Line with Secondary Axis ticked. Format the secondary axis as 0%.",
    },
    quiz: [
      {
        q: "A sparkline is:",
        options: [
          "A chart on its own sheet",
          "A tiny chart inside a single cell",
          "A trendline",
          "A pivot chart",
        ],
        answer: 1,
        explain: "Sparklines live in a cell and move with the row.",
      },
    ],
  },
  {
    slug: "dashboard-build",
    title: "Project: build an interactive dashboard",
    track: "charts",
    level: "Advanced",
    summary:
      "Combine Tables, PivotTables, slicers, pivot charts and KPI cards into one page that a manager can use without asking you anything.",
    objectives: [
      "Structure a workbook into Data, Calculation and Dashboard sheets",
      "Connect slicers to several pivots at once",
      "Build KPI cards and finish the layout professionally",
    ],
    steps: [
      {
        title: "1. Three sheets, three jobs",
        detail:
          "Data holds the raw table only. Calc holds the pivots and helper formulas. Dashboard holds nothing but visuals pointing at Calc. Never mix them.",
      },
      {
        title: "2. Make the source a Table",
        detail:
          "Ctrl + T on the raw data and name it (Table Design > Table Name). Every pivot then grows automatically.",
      },
      {
        title: "3. Build the pivots you need",
        detail:
          "One pivot per visual: revenue by month, revenue by region, top 10 customers, orders by status. Keep them on the Calc sheet, well apart from each other.",
      },
      {
        title: "4. Turn them into pivot charts",
        detail:
          "Select a pivot, PivotTable Analyze > PivotChart. Hide the field buttons via the chart's right-click menu for a clean look, then cut and paste the chart onto the Dashboard.",
      },
      {
        title: "5. Add slicers and a timeline",
        detail:
          "PivotTable Analyze > Insert Slicer for Region and Status, Insert Timeline for the date field.",
      },
      {
        title: "6. Connect one slicer to every pivot",
        detail:
          "Right-click the slicer > Report Connections and tick every pivot that shares the same source. Now one click filters the whole dashboard.",
      },
      {
        title: "7. KPI cards",
        detail:
          "Use GETPIVOTDATA or a CUBEVALUE-free approach: simple SUMIFS from the Table into four large cells, formatted big and bold with a coloured fill.",
        formula: '=SUMIFS(Sales[Amount];Sales[Region];$B$1)',
      },
      {
        title: "8. Add a variance indicator",
        detail:
          "Conditional Formatting > Icon Sets on the change column, or a simple formula label.",
        formula: '=IF(B5>=C5;"▲ on target";"▼ below target")',
      },
      {
        title: "9. Finish the layout",
        detail:
          "View > untick Gridlines, Headings and Formula Bar. Align every object with Page Layout > Align > Snap to Grid. Use one accent colour and grey for everything else.",
      },
      {
        title: "10. Protect and hand over",
        detail:
          "Hide the Data and Calc sheets, protect the Dashboard sheet allowing 'Use PivotTable reports' and slicer use, and add a Refresh button (Data > Refresh All, or a one-line macro).",
        formula: "Sub RefreshAll_Click()\n    ThisWorkbook.RefreshAll\nEnd Sub",
      },
    ],
    practice: {
      task: "A manager wants revenue by month and by region, filterable by status, on one page.",
      answer:
        "Table the data, build two pivots on Calc, create two pivot charts, insert a Status slicer, use Report Connections to link it to both pivots, paste the charts on Dashboard and hide the other sheets.",
    },
    quiz: [
      {
        q: "How do you make one slicer filter several PivotTables?",
        options: [
          "Copy the slicer",
          "Right-click the slicer > Report Connections",
          "Group the sheets",
          "It is not possible",
        ],
        answer: 1,
        explain: "Report Connections links a slicer to every pivot sharing the same data source.",
      },
    ],
  },
];
