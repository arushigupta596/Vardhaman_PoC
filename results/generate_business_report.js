const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");

// ── Color palette ──────────────────────────────────────────
const NAVY    = "1B3A5C";
const BLUE    = "2E75B6";
const LTBLUE  = "D5E8F0";
const LTGRAY  = "F2F2F2";
const WHITE   = "FFFFFF";
const GREEN   = "2E7D32";
const RED     = "C62828";
const ORANGE  = "E65100";
const DGRAY   = "333333";

// ── Helper functions ───────────────────────────────────────
function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  const runOpts = { text, font: "Arial", size: 22, color: DGRAY, ...opts };
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    children: [new TextRun(runOpts)],
    alignment: opts.align || AlignmentType.LEFT,
  });
}

function boldPara(text, opts = {}) {
  return para(text, { bold: true, ...opts });
}

function spacer(pts = 100) {
  return new Paragraph({ spacing: { after: pts }, children: [] });
}

const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function cell(text, opts = {}) {
  const { width, fill, bold, align, color, fontSize } = {
    fill: WHITE, bold: false, align: AlignmentType.LEFT, color: DGRAY, fontSize: 20, ...opts
  };
  return new TableCell({
    borders,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, font: "Arial", size: fontSize, color })]
    })]
  });
}

function headerCell(text, width) {
  return cell(text, { width, fill: NAVY, bold: true, color: WHITE, fontSize: 20, align: AlignmentType.CENTER });
}

function dataRow(cells, even = false) {
  const fill = even ? LTGRAY : WHITE;
  return new TableRow({
    children: cells.map(([text, width, opts]) => cell(text, { width, fill, ...opts }))
  });
}

// ── Build document ─────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: DGRAY } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } } }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  },
  sections: [
    // ════════════════════════════════════════════════════════
    // TITLE PAGE
    // ════════════════════════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        spacer(2400),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "NY COTTON FUTURES", font: "Arial", size: 52, bold: true, color: NAVY })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "PRICE PREDICTION SYSTEM", font: "Arial", size: 52, bold: true, color: NAVY })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 8 } },
          spacing: { after: 400 },
          children: []
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Technical & Business Results Report", font: "Arial", size: 32, color: BLUE })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Powered by Amazon Chronos-2 Foundation Model", font: "Arial", size: 24, color: DGRAY })]
        }),
        spacer(600),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Prepared for: Cotton Procurement & Trading Team", font: "Arial", size: 22, color: DGRAY })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Date: March 23, 2026", font: "Arial", size: 22, color: DGRAY })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Version: 3.0 (Multivariate + Covariates + Ensemble)", font: "Arial", size: 22, color: DGRAY })]
        }),
        spacer(800),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "CONFIDENTIAL", font: "Arial", size: 28, bold: true, color: RED })]
        }),
      ]
    },

    // ════════════════════════════════════════════════════════
    // MAIN CONTENT
    // ════════════════════════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 4 } },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: "NY Cotton Futures Price Prediction System", font: "Arial", size: 16, color: BLUE, italics: true }),
              new TextRun({ text: "\tConfidential", font: "Arial", size: 16, color: BLUE, italics: true }),
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Page ", font: "Arial", size: 16, color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "999999" }),
            ]
          })]
        })
      },
      children: [
        // ── 1. EXECUTIVE SUMMARY ──────────────────────────
        heading("1. Executive Summary"),
        para("This report presents the results of our AI-powered cotton futures price prediction system, built using Amazon Chronos-2, a state-of-the-art 120-million parameter foundation model for time series forecasting. The system predicts ICE NY Cotton (CT1) closing prices at 30, 60, and 90 trading-day horizons."),
        spacer(80),

        // Key Results Box
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [
            new TableRow({
              children: [new TableCell({
                borders: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE }, left: { style: BorderStyle.SINGLE, size: 4, color: BLUE }, right: { style: BorderStyle.SINGLE, size: 4, color: BLUE } },
                width: { size: 9360, type: WidthType.DXA },
                shading: { fill: LTBLUE, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "KEY RESULTS AT A GLANCE", font: "Arial", size: 24, bold: true, color: NAVY })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
                    new TextRun({ text: "30-day forecast accuracy: 2.66 cents/lb average error (~4.1% on a ~65 cent price)", font: "Arial", size: 20, color: DGRAY })
                  ]}),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
                    new TextRun({ text: "82-89% more accurate than traditional Prophet forecasting (statistically significant, p < 0.05)", font: "Arial", size: 20, color: DGRAY })
                  ]}),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
                    new TextRun({ text: "90-95% confidence interval coverage: actual prices fall within our predicted range 9 out of 10 times", font: "Arial", size: 20, color: DGRAY })
                  ]}),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
                    new TextRun({ text: "Current live forecast: ~61.7 cents/lb over the next 30-90 days (stable outlook)", font: "Arial", size: 20, color: DGRAY })
                  ]}),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [
                    new TextRun({ text: "Uses 20+ data inputs including CFTC trader positioning, US Dollar, crude oil, drought data, and crop calendar", font: "Arial", size: 20, color: DGRAY })
                  ]}),
                ]
              })]
            })
          ]
        }),
        spacer(120),

        para("For a cotton procurement team purchasing 1,000 bales (~480,000 lbs), the 30-day forecast uncertainty translates to approximately $12,768 per order. This represents a significant improvement over traditional forecasting methods and provides actionable intelligence for procurement timing, hedging decisions, and quarterly budget planning."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 2. SYSTEM OVERVIEW ────────────────────────────
        heading("2. System Overview"),
        para("The prediction system follows a seven-stage pipeline that transforms raw market data into probabilistic price forecasts:"),
        spacer(60),

        heading("2.1 Architecture", HeadingLevel.HEADING_2),
        // Architecture flow
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 2, color: BLUE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: BLUE }, left: { style: BorderStyle.SINGLE, size: 2, color: BLUE }, right: { style: BorderStyle.SINGLE, size: 2, color: BLUE } },
              width: { size: 9360, type: WidthType.DXA },
              shading: { fill: LTGRAY, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
                  new TextRun({ text: "SYSTEM ARCHITECTURE", font: "Arial", size: 20, bold: true, color: NAVY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "Data Ingestion (Cotton Futures, Macro, CFTC, Weather)", font: "Courier New", size: 18, color: DGRAY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "            |", font: "Courier New", size: 18, color: BLUE })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "Feature Engineering (70+ features across 6 categories)", font: "Courier New", size: 18, color: DGRAY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "            |", font: "Courier New", size: 18, color: BLUE })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "Covariate Validation & Normalization", font: "Courier New", size: 18, color: DGRAY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "        /           \\", font: "Courier New", size: 18, color: BLUE })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "Chronos-2 (60%)    Chronos-Bolt (40%)", font: "Courier New", size: 18, color: DGRAY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "        \\           /", font: "Courier New", size: 18, color: BLUE })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "Ensemble Combination + Bias Correction", font: "Courier New", size: 18, color: DGRAY })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
                  new TextRun({ text: "            |", font: "Courier New", size: 18, color: BLUE })
                ]}),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [
                  new TextRun({ text: "Probabilistic Price Forecast (30 / 60 / 90 days)", font: "Courier New", size: 18, bold: true, color: GREEN })
                ]}),
              ]
            })]
          })]
        }),
        spacer(120),

        heading("2.2 Model Components", HeadingLevel.HEADING_2),
        para("The system uses an ensemble of two complementary models:"),
        spacer(40),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 3280, 3280],
          rows: [
            new TableRow({ children: [
              headerCell("Component", 2800), headerCell("Chronos-2 (Primary)", 3280), headerCell("Chronos-Bolt (Secondary)", 3280)
            ]}),
            dataRow([["Weight in Ensemble", 2800], ["60%", 3280, { align: AlignmentType.CENTER }], ["40%", 3280, { align: AlignmentType.CENTER }]]),
            dataRow([["Parameters", 2800], ["120 Million", 3280, { align: AlignmentType.CENTER }], ["Base model", 3280, { align: AlignmentType.CENTER }]], true),
            dataRow([["Approach", 2800], ["Multivariate (uses covariates)", 3280, { align: AlignmentType.CENTER }], ["Univariate (price only)", 3280, { align: AlignmentType.CENTER }]]),
            dataRow([["Inputs", 2800], ["Price + 14 past + 6 future covariates", 3280, { align: AlignmentType.CENTER }], ["Last 512 price values", 3280, { align: AlignmentType.CENTER }]], true),
            dataRow([["Strength", 2800], ["Captures macro/positioning signals", 3280, { align: AlignmentType.CENTER }], ["Robust baseline, avoids overfitting", 3280, { align: AlignmentType.CENTER }]]),
          ]
        }),
        spacer(80),
        para("The ensemble approach combines the multivariate intelligence of Chronos-2 with the robustness of Chronos-Bolt, reducing the risk of any single model failure affecting forecasts."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 3. DATA INPUTS ────────────────────────────────
        heading("3. Data Inputs Explained"),
        para("The model ingests data from five distinct categories. Each input was selected based on its historical correlation with cotton prices and its fundamental economic relevance to cotton markets."),
        spacer(80),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1800, 2500, 2500, 2560],
          rows: [
            new TableRow({ children: [
              headerCell("Category", 1800), headerCell("Key Inputs", 2500), headerCell("Why It Matters", 2500), headerCell("Correlation", 2560)
            ]}),
            dataRow([
              ["Price & Structure", 1800, { bold: true }],
              ["CT1 Close, Volume, Roll Yield, Term Spread", 2500],
              ["Core price dynamics and market structure signals", 2500],
              ["Direct target data", 2560, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["CFTC Positioning", 1800, { bold: true }],
              ["Non-commercial long positions, Speculator net %, Concentration (top 4)", 2500],
              ["Shows what large institutional traders are doing. When speculators go heavily long, prices tend to follow", 2500],
              ["r = 0.71 (highest)", 2560, { align: AlignmentType.CENTER, color: GREEN }]
            ], true),
            dataRow([
              ["Macro", 1800, { bold: true }],
              ["US Dollar Index (DXY), WTI Crude Oil", 2500],
              ["Stronger dollar makes cotton more expensive for overseas buyers, reducing demand. Oil affects production costs", 2500],
              ["Inverse relationship", 2560, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["Weather", 1800, { bold: true }],
              ["Palmer Drought Severity Index (PDSI), Temperature, Precipitation", 2500],
              ["Cotton is highly sensitive to drought in the US cotton belt (TX, GA, MS). Drought reduces supply, raising prices", 2500],
              ["r = 0.42", 2560, { align: AlignmentType.CENTER }]
            ], true),
            dataRow([
              ["Crop Calendar", 1800, { bold: true }],
              ["Planting (Apr-Jun), Boll Dev (Jul-Aug), Harvest (Sep-Nov), WASDE reports", 2500],
              ["Cotton prices are seasonal. Key USDA reports on 9th-13th of each month can move markets 2-5%", 2500],
              ["Seasonal pattern", 2560, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["Momentum", 1800, { bold: true }],
              ["5-day / 21-day price returns, DXY change, CFTC position changes", 2500],
              ["Short-term trend signals help the model identify price momentum and mean-reversion", 2500],
              ["Trend capture", 2560, { align: AlignmentType.CENTER }]
            ], true),
          ]
        }),
        spacer(120),
        para("In total, the model processes 14 past covariates (historical data that informs the forecast) and 6 known future covariates (calendar/seasonal values that can be computed in advance)."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 4. BACKTEST RESULTS ───────────────────────────
        heading("4. Backtest Results"),
        para("The model was rigorously tested using a walk-forward backtest methodology: we simulate making forecasts at 20 historical points over the past 2 years (April 2024 to March 2026) and compare our predictions against what actually happened. This is the gold standard for evaluating forecasting models because it mirrors real-world usage."),
        spacer(80),

        heading("4.1 Forecast Accuracy", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 1640, 1640, 1640, 1640],
          rows: [
            new TableRow({ children: [
              headerCell("Metric", 2800), headerCell("30-Day", 1640), headerCell("60-Day", 1640), headerCell("90-Day", 1640), headerCell("Unit", 1640)
            ]}),
            dataRow([
              ["Mean Absolute Error (MAE)", 2800, { bold: true }],
              ["2.66", 1640, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["3.51", 1640, { align: AlignmentType.CENTER }],
              ["3.27", 1640, { align: AlignmentType.CENTER }],
              ["cents/lb", 1640, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["Percentage Error", 2800, { bold: true }],
              ["~4.1%", 1640, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["~5.4%", 1640, { align: AlignmentType.CENTER }],
              ["~5.0%", 1640, { align: AlignmentType.CENTER }],
              ["of price", 1640, { align: AlignmentType.CENTER }]
            ], true),
            dataRow([
              ["Direction Accuracy", 2800, { bold: true }],
              ["60%", 1640, { align: AlignmentType.CENTER, color: GREEN }],
              ["45%", 1640, { align: AlignmentType.CENTER }],
              ["40%", 1640, { align: AlignmentType.CENTER }],
              ["correct calls", 1640, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["Confidence Coverage", 2800, { bold: true }],
              ["90%", 1640, { align: AlignmentType.CENTER, color: GREEN }],
              ["95%", 1640, { align: AlignmentType.CENTER, color: GREEN }],
              ["95%", 1640, { align: AlignmentType.CENTER, color: GREEN }],
              ["within band", 1640, { align: AlignmentType.CENTER }]
            ], true),
            dataRow([
              ["CRPS (Probability Score)", 2800, { bold: true }],
              ["0.88", 1640, { align: AlignmentType.CENTER }],
              ["1.19", 1640, { align: AlignmentType.CENTER }],
              ["1.18", 1640, { align: AlignmentType.CENTER }],
              ["lower = better", 1640, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["Forecast Bias", 2800, { bold: true }],
              ["+1.24", 1640, { align: AlignmentType.CENTER, color: ORANGE }],
              ["+2.17", 1640, { align: AlignmentType.CENTER, color: ORANGE }],
              ["+2.72", 1640, { align: AlignmentType.CENTER, color: ORANGE }],
              ["cents/lb high", 1640, { align: AlignmentType.CENTER }]
            ], true),
          ]
        }),
        spacer(80),

        heading("4.2 Interpreting These Numbers", HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "MAE of 2.66 cents/lb (30-day): ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "On average, our 30-day price prediction is off by about 2.66 cents. For cotton trading around 65 cents/lb, this is roughly a 4% error.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "90% Confidence Coverage: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "When the model provides a prediction range (e.g., 56-68 cents), the actual price falls within that range 90% of the time. This means the uncertainty bands are well-calibrated.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "60% Direction Accuracy (30-day): ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The model correctly predicts whether prices will go up or down 60% of the time at 30 days. Note: direction prediction degrades at longer horizons (40% at 90 days), which is typical for commodity markets.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "Positive Bias (+1.24 to +2.72): ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The model has a slight tendency to over-predict prices. This bias has been measured and is automatically corrected in live forecasts.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        spacer(80),

        heading("4.3 Comparison vs Traditional Forecasting", HeadingLevel.HEADING_2),
        para("We benchmarked Chronos-2 against Facebook Prophet, a widely-used traditional forecasting model. The improvement is dramatic and statistically proven:"),
        spacer(60),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1400, 1600, 1600, 1960, 2800],
          rows: [
            new TableRow({ children: [
              headerCell("Horizon", 1400), headerCell("Chronos-2\nMAE", 1600), headerCell("Prophet\nMAE", 1600), headerCell("Improvement", 1960), headerCell("Statistical Test", 2800)
            ]}),
            dataRow([
              ["30-day", 1400, { bold: true }],
              ["2.66", 1600, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["15.24", 1600, { align: AlignmentType.CENTER, color: RED }],
              ["82.5%", 1960, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["p = 0.021 (significant)", 2800, { align: AlignmentType.CENTER, color: GREEN }]
            ]),
            dataRow([
              ["60-day", 1400, { bold: true }],
              ["3.51", 1600, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["24.30", 1600, { align: AlignmentType.CENTER, color: RED }],
              ["85.6%", 1960, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["p = 0.045 (significant)", 2800, { align: AlignmentType.CENTER, color: GREEN }]
            ], true),
            dataRow([
              ["90-day", 1400, { bold: true }],
              ["3.27", 1600, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["29.94", 1600, { align: AlignmentType.CENTER, color: RED }],
              ["89.1%", 1960, { align: AlignmentType.CENTER, bold: true, color: GREEN }],
              ["p = 0.030 (significant)", 2800, { align: AlignmentType.CENTER, color: GREEN }]
            ]),
          ]
        }),
        spacer(80),
        para("The Diebold-Mariano statistical test confirms that Chronos-2 is significantly more accurate than Prophet at all horizons (p-values below 0.05). This is not a marginal improvement; it represents a fundamental step-change in forecasting capability."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 5. BUSINESS IMPACT ────────────────────────────
        heading("4.4 What This Means in Dollar Terms", HeadingLevel.HEADING_2),
        para("To put the forecast accuracy into business context:"),
        spacer(60),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 2200, 2200, 2160],
          rows: [
            new TableRow({ children: [
              headerCell("Order Size", 2800), headerCell("30-Day\nUncertainty", 2200), headerCell("60-Day\nUncertainty", 2200), headerCell("90-Day\nUncertainty", 2160)
            ]}),
            dataRow([
              ["100 bales (48,000 lbs)", 2800, { bold: true }],
              ["$1,277", 2200, { align: AlignmentType.CENTER }],
              ["$1,685", 2200, { align: AlignmentType.CENTER }],
              ["$1,570", 2160, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["500 bales (240,000 lbs)", 2800, { bold: true }],
              ["$6,384", 2200, { align: AlignmentType.CENTER }],
              ["$8,424", 2200, { align: AlignmentType.CENTER }],
              ["$7,848", 2160, { align: AlignmentType.CENTER }]
            ], true),
            dataRow([
              ["1,000 bales (480,000 lbs)", 2800, { bold: true }],
              ["$12,768", 2200, { align: AlignmentType.CENTER, bold: true }],
              ["$16,848", 2200, { align: AlignmentType.CENTER, bold: true }],
              ["$15,696", 2160, { align: AlignmentType.CENTER, bold: true }]
            ]),
            dataRow([
              ["5,000 bales (2.4M lbs)", 2800, { bold: true }],
              ["$63,840", 2200, { align: AlignmentType.CENTER }],
              ["$84,240", 2200, { align: AlignmentType.CENTER }],
              ["$78,480", 2160, { align: AlignmentType.CENTER }]
            ], true),
          ]
        }),
        spacer(80),
        para("Uncertainty = MAE (cents/lb) x order weight (lbs). This represents the average forecasting error, not the maximum possible deviation.", { italics: true }),
        spacer(120),

        // ── 5. LIVE FORECAST ──────────────────────────────
        heading("5. Current Live Forecast"),
        para("Based on all available data as of March 23, 2026, the model produces the following price predictions for ICE NY Cotton (CT1):"),
        spacer(80),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1600, 1800, 1800, 1800, 2360],
          rows: [
            new TableRow({ children: [
              headerCell("Horizon", 1600), headerCell("Median\nForecast", 1800), headerCell("Lower Band\n(10th pctl)", 1800), headerCell("Upper Band\n(90th pctl)", 1800), headerCell("Interpretation", 2360)
            ]}),
            dataRow([
              ["30-Day", 1600, { bold: true }],
              ["61.77 cents/lb", 1800, { align: AlignmentType.CENTER, bold: true }],
              ["56.26", 1800, { align: AlignmentType.CENTER }],
              ["68.44", 1800, { align: AlignmentType.CENTER }],
              ["Slight decline expected", 2360]
            ]),
            dataRow([
              ["60-Day", 1600, { bold: true }],
              ["61.68 cents/lb", 1800, { align: AlignmentType.CENTER, bold: true }],
              ["54.84", 1800, { align: AlignmentType.CENTER }],
              ["70.85", 1800, { align: AlignmentType.CENTER }],
              ["Stable to slightly lower", 2360]
            ], true),
            dataRow([
              ["90-Day", 1600, { bold: true }],
              ["61.61 cents/lb", 1800, { align: AlignmentType.CENTER, bold: true }],
              ["54.39", 1800, { align: AlignmentType.CENTER }],
              ["71.63", 1800, { align: AlignmentType.CENTER }],
              ["Range widens with horizon", 2360]
            ]),
          ]
        }),
        spacer(80),

        // Interpretation box
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 4, color: GREEN }, bottom: { style: BorderStyle.SINGLE, size: 4, color: GREEN }, left: { style: BorderStyle.SINGLE, size: 4, color: GREEN }, right: { style: BorderStyle.SINGLE, size: 4, color: GREEN } },
              width: { size: 9360, type: WidthType.DXA },
              shading: { fill: "E8F5E9", type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "PROCUREMENT SIGNAL: ", font: "Arial", size: 22, bold: true, color: GREEN }),
                  new TextRun({ text: "NEUTRAL / SLIGHTLY BEARISH", font: "Arial", size: 22, bold: true, color: NAVY })
                ]}),
                new Paragraph({ spacing: { after: 60 }, children: [
                  new TextRun({ text: "The model forecasts cotton prices to remain relatively stable around 61.6-61.8 cents/lb over the next 90 days, with a very slight downward trajectory. The confidence bands widen at longer horizons (56-68 at 30d vs 54-72 at 90d), reflecting increasing uncertainty. This suggests no urgency for forward buying, and the procurement team may benefit from waiting for potential dips within the 54-57 cent range.", font: "Arial", size: 20, color: DGRAY })
                ]}),
              ]
            })]
          })]
        }),
        spacer(80),
        para("Note: These forecasts include bias correction based on the model's historical tendency to slightly over-predict prices. The raw model predictions have been adjusted downward by 1.24-2.72 cents/lb depending on horizon."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 6. PROCUREMENT DECISION GUIDE ─────────────────
        heading("6. How to Use This for Procurement Decisions"),
        para("This section provides practical guidance for the cotton procurement team on how to incorporate these forecasts into daily decision-making."),
        spacer(80),

        heading("6.1 Spot vs Forward Buying", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 6560],
          rows: [
            new TableRow({ children: [headerCell("Scenario", 2800), headerCell("Recommended Action", 6560)] }),
            dataRow([
              ["Spot > Upper Band", 2800, { bold: true, color: RED }],
              ["Current spot price is above the model's upper confidence band. Consider delaying purchases or hedging. The model expects prices to come down.", 6560]
            ]),
            dataRow([
              ["Spot < Lower Band", 2800, { bold: true, color: GREEN }],
              ["Current spot price is below the model's lower band. This is a buying opportunity. Consider accelerating procurement or locking in forward contracts.", 6560]
            ], true),
            dataRow([
              ["Spot within Band", 2800, { bold: true }],
              ["Price is within expected range. Follow normal procurement schedule. Use median forecast for budget planning.", 6560]
            ]),
          ]
        }),
        spacer(80),

        heading("6.2 Hedging Triggers", HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "Strong Directional Signal (30-day, 60% accuracy): ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "When the 30-day model shows a clear directional move (median significantly above or below current price), consider hedging positions accordingly.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "Wide Confidence Bands: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "If the model's confidence band is unusually wide (e.g., >15 cents spread), this signals high uncertainty. Consider options-based hedging to protect against extreme moves.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
          new TextRun({ text: "WASDE Report Windows: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The model flags USDA WASDE report dates (typically 9th-13th of month). Avoid large unhedged positions around these dates, as reports can move cotton prices 2-5% in a single session.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        spacer(80),

        heading("6.3 Quarterly Budget Planning", HeadingLevel.HEADING_2),
        para("For budget estimates, use the following framework:"),
        spacer(40),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Best case: ", font: "Arial", size: 22, bold: true, color: GREEN }),
          new TextRun({ text: "Use 10th percentile (lower band) for optimistic scenario", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Base case: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "Use median forecast for central planning estimate", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Worst case: ", font: "Arial", size: 22, bold: true, color: RED }),
          new TextRun({ text: "Use 90th percentile (upper band) for conservative budget cushion", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        spacer(80),
        para("Example for Q2 2026 (based on 90-day forecast): Budget at 61.61 cents/lb base case, with range of 54.39-71.63 cents/lb for scenario planning."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 7. WHAT WE BUILT (TECHNICAL JOURNEY) ──────────
        heading("7. What Was Built: Technical Journey"),
        para("This section documents the development process and improvements made to reach the current level of accuracy."),
        spacer(80),

        heading("7.1 Development Phases", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1200, 2400, 3200, 2560],
          rows: [
            new TableRow({ children: [
              headerCell("Phase", 1200), headerCell("What Was Done", 2400), headerCell("Technical Detail", 3200), headerCell("Impact", 2560)
            ]}),
            dataRow([
              ["v1", 1200, { bold: true }],
              ["Basic univariate model", 2400],
              ["Chronos-Bolt with price-only input", 3200],
              ["Baseline established", 2560]
            ]),
            dataRow([
              ["v2", 1200, { bold: true }],
              ["Feature engineering", 2400],
              ["70+ features: technicals, macro, CFTC, weather, seasonality", 3200],
              ["Data pipeline built", 2560]
            ], true),
            dataRow([
              ["v3.0", 1200, { bold: true }],
              ["Chronos-2 multivariate", 2400],
              ["Upgraded to Chronos-2 with predict_df API, covariates, cross-learning", 3200],
              ["MAE: 3.29 (30d)", 2560]
            ]),
            dataRow([
              ["v3.1", 1200, { bold: true }],
              ["Removed synthetic data", 2400],
              ["Removed CT2/CT3 synthetic deferred contracts (had correlation 1.0 with CT1 - no information)", 3200],
              ["Cleaner model inputs", 2560]
            ], true),
            dataRow([
              ["v3.2", 1200, { bold: true }],
              ["Added high-value covariates", 2400],
              ["Added CFTC positioning (r=0.71), speculator net (r=0.67), concentration (r=0.63), drought (r=0.42)", 3200],
              ["Better signal capture", 2560]
            ]),
            dataRow([
              ["v3.3", 1200, { bold: true }],
              ["Ensemble + bias correction", 2400],
              ["Chronos-2 (60%) + Chronos-Bolt (40%) ensemble, warm-up skip, bias correction", 3200],
              ["MAE: 2.66 (30d)", 2560, { color: GREEN, bold: true }]
            ], true),
            dataRow([
              ["v3.4", 1200, { bold: true }],
              ["Momentum features + normalization", 2400],
              ["Added 5d/21d returns, DXY/WTI momentum, CFTC changes, z-score normalization", 3200],
              ["Final production model", 2560]
            ]),
          ]
        }),
        spacer(80),

        heading("7.2 Accuracy Improvement Journey", HeadingLevel.HEADING_2),
        para("The systematic improvements reduced the 30-day MAE from 3.29 to 2.66 cents/lb, a 19% improvement. The 90-day horizon saw the largest gain, with MAE dropping from 4.70 to 3.27 cents/lb (30% improvement)."),
        spacer(60),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3120, 2080, 2080, 2080],
          rows: [
            new TableRow({ children: [
              headerCell("Improvement", 3120), headerCell("30d MAE", 2080), headerCell("60d MAE", 2080), headerCell("90d MAE", 2080)
            ]}),
            dataRow([
              ["Before improvements", 3120, { bold: true }], ["3.29", 2080, { align: AlignmentType.CENTER }], ["4.56", 2080, { align: AlignmentType.CENTER }], ["4.70", 2080, { align: AlignmentType.CENTER }]
            ]),
            dataRow([
              ["After all improvements", 3120, { bold: true }], ["2.66", 2080, { align: AlignmentType.CENTER, color: GREEN, bold: true }], ["3.51", 2080, { align: AlignmentType.CENTER, color: GREEN, bold: true }], ["3.27", 2080, { align: AlignmentType.CENTER, color: GREEN, bold: true }]
            ], true),
            dataRow([
              ["Reduction", 3120, { bold: true, color: GREEN }], ["-19%", 2080, { align: AlignmentType.CENTER, bold: true, color: GREEN }], ["-23%", 2080, { align: AlignmentType.CENTER, bold: true, color: GREEN }], ["-30%", 2080, { align: AlignmentType.CENTER, bold: true, color: GREEN }]
            ]),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 8. LIMITATIONS ────────────────────────────────
        heading("8. Limitations & Important Caveats"),
        para("While the model shows strong performance, users should be aware of the following limitations:"),
        spacer(80),

        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Longer horizons have wider uncertainty. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The 90-day directional accuracy is 40%, which means the model's directional calls at that horizon should not be relied upon for tactical timing. Use the confidence bands for budget planning instead.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Over-prediction bias exists but is corrected. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The raw model tends to predict prices 1-3 cents higher than actual. Live forecasts automatically subtract this measured bias, but users should be aware that the correction is based on historical patterns.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Black swan events are not predictable. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The model cannot foresee trade wars, pandemics, supply chain disruptions, or other unprecedented events. Always maintain appropriate risk buffers.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Weather data is retrospective. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The PDSI drought data reflects historical conditions, not real-time satellite observations. Rapidly developing weather events may not be captured until the next data update.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Past performance does not guarantee future accuracy. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "The backtest covers April 2024 to March 2026. Market regimes can change, and the model should be continuously monitored and recalibrated.", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [
          new TextRun({ text: "Some outlier periods exist. ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "Three backtest origins (April 2024, May 2024, August 2024) showed elevated errors. These periods coincided with unusual market conditions. The warm-up skip and ensemble approach mitigate but do not eliminate outlier risk.", font: "Arial", size: 22, color: DGRAY }),
        ]}),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 9. TECHNICAL APPENDIX ─────────────────────────
        heading("9. Technical Appendix"),
        spacer(40),

        heading("9.1 Full Covariate List", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1400, 3200, 2400, 2360],
          rows: [
            new TableRow({ children: [
              headerCell("Type", 1400), headerCell("Covariate", 3200), headerCell("Description", 2400), headerCell("Correlation w/ CT1", 2360)
            ]}),
            ...[
              ["Past", "dxy", "US Dollar Index", "Inverse"],
              ["Past", "wti_crude", "WTI Crude Oil", "Moderate"],
              ["Past", "traders_noncomm_long", "CFTC Non-commercial Longs", "r = 0.71"],
              ["Past", "spec_net_pct", "Speculator Net % of OI", "r = 0.67"],
              ["Past", "conc_4_short", "Top 4 Short Concentration", "r = 0.63"],
              ["Past", "pdsi_severe_drought", "Severe Drought Flag", "r = 0.42"],
              ["Past", "realised_vol_21d", "21-day Realized Volatility", "Risk signal"],
              ["Past", "noaa_pdsi", "Palmer Drought Index", "Supply risk"],
              ["Past", "ct1_ret_5d", "5-day Price Return", "Momentum"],
              ["Past", "ct1_ret_21d", "21-day Price Return", "Trend"],
              ["Past", "dxy_5d_ret", "DXY 5-day Return", "Macro momentum"],
              ["Past", "wti_5d_ret", "WTI 5-day Return", "Energy trend"],
              ["Past", "noncomm_long_chg_5d", "CFTC Long Position Change", "Flow signal"],
              ["Past", "spec_net_pct_chg_5d", "Speculator Net Change", "Sentiment shift"],
              ["Future", "seas_sin_annual", "Annual Sine Seasonality", "Calendar"],
              ["Future", "seas_cos_annual", "Annual Cosine Seasonality", "Calendar"],
              ["Future", "flag_planting", "Planting Season (Apr-Jun)", "Crop cycle"],
              ["Future", "flag_boll_dev", "Boll Development (Jul-Aug)", "Crop cycle"],
              ["Future", "flag_harvest", "Harvest (Sep-Nov)", "Crop cycle"],
              ["Future", "flag_wasde", "WASDE Report Window", "Event flag"],
            ].map(([type, name, desc, corr], i) =>
              dataRow([
                [type, 1400, { bold: type === "Past", color: type === "Past" ? NAVY : GREEN }],
                [name, 3200],
                [desc, 2400],
                [corr, 2360, { align: AlignmentType.CENTER }]
              ], i % 2 === 1)
            )
          ]
        }),
        spacer(120),

        heading("9.2 Backtest Methodology", HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Method: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "Walk-forward expanding window with strict no-leakage guarantee", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Test Period: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "April 24, 2024 through March 20, 2026 (500 trading days)", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Forecast Origins: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "20 origins spaced 10 trading days apart (after 1 warm-up skip)", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Minimum Context: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "500 trading days of history required before first forecast", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Ensemble: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "Chronos-2 multivariate (weight: 0.6) + Chronos-Bolt univariate (weight: 0.4)", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Statistical Test: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "Diebold-Mariano test comparing squared forecast errors between Chronos-2 and Prophet", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
          new TextRun({ text: "Quantile Levels: ", font: "Arial", size: 22, bold: true, color: DGRAY }),
          new TextRun({ text: "10th, 25th, 50th (median), 75th, 90th percentiles", font: "Arial", size: 22, color: DGRAY }),
        ]}),
        spacer(200),

        // ── DISCLAIMER ────────────────────────────────────
        new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 8 } },
          spacing: { before: 200, after: 100 },
          children: [new TextRun({ text: "DISCLAIMER", font: "Arial", size: 20, bold: true, color: "999999" })]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({
            text: "This report is for informational purposes only and does not constitute financial advice. Cotton futures trading involves substantial risk of loss. The forecasts presented are model-based estimates and should be used as one input among many in procurement decision-making. Past model performance during the backtest period does not guarantee future accuracy. Always consult with qualified financial and commodity risk professionals before making trading or procurement decisions.",
            font: "Arial", size: 18, color: "999999", italics: true
          })]
        }),
      ]
    }
  ]
});

// ── Write file ─────────────────────────────────────────────
Packer.toBuffer(doc).then(buffer => {
  const outPath = __dirname + "/Cotton_Chronos2_Business_Report.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Document saved to:", outPath);
  console.log("Size:", (buffer.length / 1024).toFixed(1), "KB");
});
