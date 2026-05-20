const express = require("express")
const cors = require("cors")
const PptxGenJS = require("pptxgenjs")

const app = express()

app.use(cors())
app.use(express.json())

app.use("/reports", express.static("reports"))

// --------------------------------------------------
// DESIGN TOKENS  — warm cream + deep red + amber brown
// --------------------------------------------------

const C = {
  BG: "F6F1E7",   // warm cream (page background)
  BG2: "EEE4D2",  // slightly deeper cream
  BG3: "FFFFFF",  // card white
  BG4: "F0E8D8",  // light card / header row
  ACCENT1: "8B0000", // deep red (primary)
  ACCENT2: "C21807", // bright red (gradient end)
  ACCENT3: "7A3D1B", // amber brown
  ACCENT4: "A0522D", // sienna
  WHITE: "FFFFFF",
  LIGHT: "2A1E17",   // dark brown text
  MUTED: "7A3D1B",   // muted amber/brown
  DARK_TXT: "2A1E17",
  CHART: ["8B0000", "C21807", "7A3D1B", "A0522D", "661111", "D2691E"],
}

const FONT_H = "Trebuchet MS"
const FONT_B = "Calibri"

// --------------------------------------------------
// HELPERS
// --------------------------------------------------

function safe(v) {
  if (v === null || v === undefined || v === "") return "No data available"
  if (typeof v === "string") return v
  return JSON.stringify(v)
}

function fmt(n) {
  if (n === null || n === undefined || n === "") return "N/A"
  const num = Number(n)
  if (Number.isNaN(num)) return "N/A"
  if (num >= 1_000_000) return (num / 1_000_000).toFixed(1) + "M"
  if (num >= 1_000) return (num / 1_000).toFixed(1) + "K"
  return num.toLocaleString()
}

function fmtFull(n) {
  if (n === null || n === undefined || n === "") return "N/A"
  const num = Number(n)
  if (Number.isNaN(num)) return "N/A"
  return num.toLocaleString()
}

function setupSlide(pptx, title, subtitle) {
  const slide = pptx.addSlide()
  slide.background = { color: C.BG }

  // Left accent bar — deep red
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.06, h: 7.5,
    fill: { color: C.ACCENT1 }, line: { color: C.ACCENT1 },
  })

  if (title) {
    slide.addText(title, {
      x: 0.35, y: 0.22, w: 12.3, h: 0.6,
      fontFace: FONT_H, fontSize: 28, bold: true, color: C.ACCENT1, margin: 0,
    })
    if (subtitle) {
      slide.addText(subtitle, {
        x: 0.35, y: 0.85, w: 12.3, h: 0.28,
        fontFace: FONT_B, fontSize: 13, color: C.MUTED, margin: 0,
      })
    }
  }
  return slide
}

function card(slide, pptx, x, y, w, h, color, radius) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    rectRadius: radius !== undefined ? radius : 0.1,
    fill: { color: color || C.BG3 },
    line: { color: "E0D8CC" },
  })
}

function statBox(slide, pptx, x, y, w, value, label, color) {
  card(slide, pptx, x, y, w, 1.05)
  slide.addText(String(value), {
    x: x + 0.1, y: y + 0.07, w: w - 0.2, h: 0.52,
    fontFace: FONT_H, fontSize: 26, bold: true, color: color || C.ACCENT1,
    align: "center", margin: 0,
  })
  slide.addText(label, {
    x: x + 0.1, y: y + 0.62, w: w - 0.2, h: 0.3,
    fontFace: FONT_B, fontSize: 12, color: C.MUTED, align: "center", margin: 0,
  })
}

function hLine(slide, pptx, x, y, w) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h: 0.025,
    fill: { color: "E0D8CC" }, line: { color: "E0D8CC" },
  })
}

function smartAxisMax(values) {
  const max = Math.max(...values, 0)

  if (max <= 10) return 10
  if (max <= 100) return Math.ceil(max * 1.4)
  if (max <= 1_000) return Math.ceil(max / 100) * 100 * 1.25
  if (max <= 10_000) return Math.ceil(max / 1_000) * 1_000 * 1.2
  if (max <= 100_000) return Math.ceil(max / 10_000) * 10_000 * 1.2
  if (max <= 1_000_000) return Math.ceil(max / 100_000) * 100_000 * 1.15
  return Math.ceil(max / 1_000_000) * 1_000_000 * 1.15
}

// --------------------------------------------------
// ROUTE
// --------------------------------------------------

app.post("/generate-ppt", async (req, res) => {
  try {
    console.log("STARTING PPT GENERATION")

    const data = req.body
    const companies = data.companies_analyzed || []
    const names = companies.map(c => c.channel_name)
    const reportDate = new Date().toLocaleDateString("en-US", {
      year: "numeric", month: "long", day: "numeric",
    })

    const pptx = new PptxGenJS()
    pptx.layout = "LAYOUT_WIDE"
    pptx.author = "Video Intelligence"
    pptx.company = "Competitor Analytics"
    pptx.title = "Video Competitor Intelligence"

    // ================================================
    // SLIDE 1 — COVER
    // ================================================
    {
      const slide = pptx.addSlide()
      slide.background = { color: C.BG }

      // Bold left accent stripe — deep red
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: 0.55, h: 7.5,
        fill: { color: C.ACCENT1 }, line: { color: C.ACCENT1 },
      })

      // Warm overlay block for text area
      card(slide, pptx, 0.55, 0, 12.45, 7.5, C.BG, 0)

      // Decorative dots grid (top right)
      for (let row = 0; row < 5; row++) {
        for (let col = 0; col < 8; col++) {
          slide.addShape(pptx.shapes.OVAL, {
            x: 8.5 + col * 0.45, y: 0.3 + row * 0.45, w: 0.1, h: 0.1,
            fill: { color: "E0D8CC" }, line: { color: "E0D8CC" },
          })
        }
      }

      slide.addText("VIDEO COMPETITOR\nINTELLIGENCE", {
        x: 0.9, y: 1.2, w: 8.5, h: 2.6,
        fontFace: FONT_H, fontSize: 52, bold: true, color: C.ACCENT1,
      })

      slide.addText("AI-Powered YouTube Competitor Analysis", {
        x: 0.9, y: 3.9, w: 9, h: 0.5,
        fontFace: FONT_B, fontSize: 17, color: C.MUTED,
      })

      // Channels list as pill badges
      names.forEach((name, i) => {
        const pillX = 0.9 + i * 3.1
        if (pillX + 2.8 > 13.2) return
        card(slide, pptx, pillX, 4.65, 2.9, 0.48, C.BG4, 0.24)
        slide.addText(name, {
          x: pillX + 0.1, y: 4.68, w: 2.7, h: 0.38,
          fontFace: FONT_B, fontSize: 13, bold: true, color: C.ACCENT1,
          align: "center", margin: 0,
        })
      })

      slide.addText(`Report Date: ${reportDate}`, {
        x: 0.9, y: 6.2, w: 6, h: 0.3,
        fontFace: FONT_B, fontSize: 11, color: C.MUTED,
      })
    }

    // ================================================
    // SLIDE 2 — EXECUTIVE SUMMARY
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Executive Summary",
        "Who is leading in video marketing — and why"
      )

      const summaryText = safe(data.executive_summary)
      const estimatedLines = Math.ceil(summaryText.length / 82)
      const dynamicHeight = Math.max(1.8, Math.min(2.5, estimatedLines * 0.34))

      const summaryY = 1.45

      card(slide, pptx, 0.55, summaryY, 12.2, dynamicHeight, C.BG3)

      // Top accent bar on summary box
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.55, y: summaryY,
        w: 12.2, h: 0.08,
        fill: { color: C.ACCENT1 }, line: { color: C.ACCENT1 },
      })

      slide.addText(summaryText, {
        x: 0.95, y: summaryY + 0.25,
        w: 11.4, h: dynamicHeight - 0.28,
        fontFace: FONT_B, fontSize: 16,
        color: C.LIGHT,
        align: "left", valign: "mid",
      })

      const kpis = [
        ["Subscribers", data?.leaderboards?.highest_subscribers?.company || "N/A"],
        ["Engagement", data?.leaderboards?.highest_engagement?.company || "N/A"],
        ["Avg Views", data?.leaderboards?.highest_average_views?.company || "N/A"],
        ["Most Active", data?.leaderboards?.most_active_channel?.company || "N/A"],
      ]

      const cardW = 2.9
      const cardGap = 0.22
      const totalWidth = (cardW * 4) + (cardGap * 3)
      const startX = (13.33 - totalWidth) / 2
      const startY = summaryY + dynamicHeight + 0.55

      kpis.forEach(([label, value], i) => {
        const x = startX + i * (cardW + cardGap)

        card(slide, pptx, x, startY, cardW, 1.45, C.BG4)

        // Accent bar
        slide.addShape(pptx.shapes.RECTANGLE, {
          x, y: startY,
          w: 0.1, h: 1.45,
          fill: { color: C.CHART[i] }, line: { color: C.CHART[i] },
        })

        slide.addText(label, {
          x: x + 0.22, y: startY + 0.18,
          w: 2.35, h: 0.26,
          fontFace: FONT_B, fontSize: 14,
          color: C.MUTED, align: "center", margin: 0,
        })

        slide.addText(value, {
          x: x + 0.15, y: startY + 0.62,
          w: 2.45, h: 0.42,
          fontFace: FONT_H, fontSize: 22, bold: true,
          color: C.ACCENT1, align: "center", margin: 0,
        })
      })
    }

    // ================================================
    // SLIDE 3 — CHANNEL OVERVIEW COMPARISON
    // ================================================
    {
      const slide = setupSlide(pptx, "Channel Overview Comparison")

      const count = companies.length
      const totalGap = 0.18 * (count - 1)
      const colW = (12.5 - totalGap) / count
      const gap = 0.18

      companies.forEach((c, i) => {
        const x = 0.4 + i * (colW + gap)
        const accentColor = C.CHART[i % C.CHART.length]

        card(slide, pptx, x, 1.15, colW, 6.1, C.BG3)

        // Top accent bar
        slide.addShape(pptx.shapes.RECTANGLE, {
          x, y: 1.15, w: colW, h: 0.18,
          fill: { color: accentColor }, line: { color: accentColor },
        })

        slide.addText(c.channel_name, {
          x: x + 0.12, y: 1.48,
          w: colW - 0.24, h: 0.42,
          fontFace: FONT_H, fontSize: 17, bold: true,
          color: C.ACCENT1, align: "center", margin: 0,
        })

        hLine(slide, pptx, x + 0.18, 1.95, colW - 0.36)

        const rows = [
          ["Subscribers", fmt(c.subscribers)],
          ["Total Videos", fmtFull(c.total_videos)],
          ["Avg Views", fmt(c.average_views)],
          ["Engagement", `${c.engagement_rate}%`],
          ["Uploads/week", `~${Math.round((c.upload_frequency_per_day || 0) * 7)}`],
          ["Cadence", safe(c.upload_cadence)],
        ]

        rows.forEach(([label, val], ri) => {
          const ry = 2.15 + ri * 0.82

          slide.addText(label, {
            x: x + 0.18, y: ry,
            w: colW - 0.36, h: 0.26,
            fontFace: FONT_B, fontSize: 13, color: C.MUTED,
          })

          slide.addText(val, {
            x: x + 0.18, y: ry + 0.28,
            w: colW - 0.36, h: 0.34,
            fontFace: FONT_B, fontSize: 15, bold: true, color: C.LIGHT,
          })
        })
      })
    }

    // ================================================
    // SLIDE 4 — CONTENT PERFORMANCE
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Content Performance",
        "Top performing videos by views and engagement per channel"
      )

      companies.slice(0, 4).forEach((c, i) => {
        const row = Math.floor(i / 2)
        const col = i % 2
        const x = 0.35 + col * 6.4
        const y = 1.3 + row * 3.1
        const accentColor = C.CHART[i % C.CHART.length]

        card(slide, pptx, x, y, 6.1, 2.85, C.BG3)

        slide.addShape(pptx.shapes.RECTANGLE, {
          x, y, w: 0.12, h: 2.85,
          fill: { color: accentColor }, line: { color: accentColor },
        })

        slide.addText(c.channel_name, {
          x: x + 0.3, y: y + 0.12,
          w: 5.65, h: 0.35,
          fontFace: FONT_H, fontSize: 16, bold: true, color: C.ACCENT1, margin: 0,
        })

        slide.addText(fmt(c.average_views), {
          x: x + 0.3, y: y + 0.55,
          w: 2.4, h: 0.62,
          fontFace: FONT_H, fontSize: 34, bold: true, color: accentColor, margin: 0,
        })

        slide.addText("avg views/video", {
          x: x + 0.3, y: y + 1.12,
          w: 2.5, h: 0.28,
          fontFace: FONT_B, fontSize: 14, color: C.MUTED, margin: 0,
        })

        slide.addText(`${c.engagement_rate}%`, {
          x: x + 3.0, y: y + 0.55,
          w: 2.0, h: 0.62,
          fontFace: FONT_H, fontSize: 34, bold: true, color: C.ACCENT3, margin: 0,
        })

        slide.addText("engagement rate", {
          x: x + 3.0, y: y + 1.12,
          w: 2.2, h: 0.28,
          fontFace: FONT_B, fontSize: 14, color: C.MUTED, margin: 0,
        })

        hLine(slide, pptx, x + 0.3, y + 1.55, 5.5)

        const topVid = c.top_video
        if (topVid) {
          const cleanTitle = (topVid.title || "N/A").replace(/\s+/g, " ").trim()
          const titleLines = Math.ceil(cleanTitle.length / 42)
          const titleHeight = Math.min(0.75, Math.max(0.32, titleLines * 0.18))

          slide.addText(
            [
              { text: "TOP VIDEO: ", options: { bold: true, color: C.ACCENT1 } },
              { text: cleanTitle, options: { bold: false, color: C.LIGHT } },
            ],
            {
              x: x + 0.3, y: y + 1.66,
              w: 5.35, h: titleHeight,
              fontFace: FONT_B, fontSize: 13.5,
              valign: "mid", margin: 0, breakLine: false,
            }
          )

          slide.addText(
            `${fmt(topVid.views || 0)} views · ${fmtFull(topVid.likes || 0)} likes`,
            {
              x: x + 0.3, y: y + 1.66 + titleHeight + 0.08,
              w: 5.4, h: 0.28,
              fontFace: FONT_B, fontSize: 12.5, color: C.MUTED, margin: 0,
            }
          )
        }
      })
    }

    // ================================================
    // SLIDE 5 — CONTENT TOPICS & THEMES
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Content Topics & Themes",
        "What each competitor covers — and what they are missing"
      )

      const contentW = 11.8
      const startX = (13.33 - contentW) / 2
      const themesText = safe(data.content_themes)
      const themeLines = Math.ceil(themesText.length / 82)
      const themeH = Math.max(2.1, Math.min(2.8, themeLines * 0.32))
      const themeY = 1.3

      card(slide, pptx, startX, themeY, contentW, themeH, C.BG3)

      slide.addShape(pptx.shapes.RECTANGLE, {
        x: startX, y: themeY, w: contentW, h: 0.08,
        fill: { color: C.ACCENT1 }, line: { color: C.ACCENT1 },
      })

      slide.addText("CONTENT THEMES", {
        x: startX + 0.28, y: themeY + 0.18,
        w: contentW - 0.56, h: 0.28,
        fontFace: FONT_H, fontSize: 16, bold: true, color: C.ACCENT1, margin: 0,
      })

      slide.addText(themesText, {
        x: startX + 0.28, y: themeY + 0.58,
        w: contentW - 0.56, h: themeH - 0.72,
        fontFace: FONT_B, fontSize: 15, color: C.LIGHT, valign: "mid", margin: 0,
      })

      const rawMissing = safe(data.missing_opportunities)
      const missingItems = rawMissing
        .split("\n")
        .map(l => l.trim())
        .filter(l => l.startsWith("-"))
        .map(l => l.replace(/^-\s*/, ""))
        .slice(0, 5)

      const missingH = Math.max(2.3, 0.9 + (missingItems.length * 0.5))
      const missingY = themeY + themeH + 0.38

      card(slide, pptx, startX, missingY, contentW, missingH, C.BG4)

      slide.addShape(pptx.shapes.RECTANGLE, {
        x: startX, y: missingY, w: contentW, h: 0.08,
        fill: { color: C.ACCENT3 }, line: { color: C.ACCENT3 },
      })

      slide.addText("WHAT THEY ARE MISSING", {
        x: startX + 0.28, y: missingY + 0.18,
        w: contentW - 0.56, h: 0.28,
        fontFace: FONT_H, fontSize: 16, bold: true, color: C.ACCENT3, margin: 0,
      })

      missingItems.forEach((item, idx) => {
        slide.addText(`• ${item}`, {
          x: startX + 0.28, y: missingY + 0.62 + idx * 0.48,
          w: contentW - 0.56, h: 0.34,
          fontFace: FONT_B, fontSize: 14, color: C.LIGHT, margin: 0,
        })
      })
    }

    // ================================================
    // SLIDE 6 — POSTING FREQUENCY & CONSISTENCY
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Posting Frequency & Consistency",
        "Upload volume and activity levels across all channels"
      )

      slide.addChart(
        pptx.charts.BAR,
        [{
          name: "Videos Per Week",
          labels: companies.map(c => c.channel_name),
          values: companies.map(c => Math.round((c.upload_frequency_per_day || 0) * 7)),
        }],
        {
          x: 0.75, y: 1.4, w: 7.0, h: 3.75,
          barDir: "col",
          chartColors: companies.map((_, i) => C.CHART[i % C.CHART.length]),
          chartArea: { fill: { color: C.BG3 } },
          plotArea: { fill: { color: C.BG3 } },
          catAxisLabelColor: C.MUTED,
          valAxisLabelColor: C.MUTED,
          catAxisLabelFontSize: 14,
          valAxisLabelFontSize: 12,
          catGridLine: { style: "none" },
          valGridLine: { color: "E0D8CC", style: "solid", pt: 0.5 },
          showValue: true,
          dataLabelColor: C.WHITE,
          dataLabelFontSize: 14,
          showLegend: false,
          valAxisMinVal: 0,
        }
      )

      slide.addText("ACTIVITY RATING", {
        x: 8.35, y: 1.28, w: 4.7, h: 0.3,
        fontFace: FONT_H, fontSize: 14, bold: true, color: C.MUTED, margin: 0,
      })

      const cadenceColor = (c) => {
        if (c === "Very Active") return C.ACCENT1
        if (c === "Moderately Active") return C.ACCENT3
        return "C21807"
      }

      companies.forEach((c, i) => {
        const y = 1.7 + i * 1.12
        const accent = C.CHART[i % C.CHART.length]
        const uploadsPerWeek = Math.round((c.upload_frequency_per_day || 0) * 7)

        card(slide, pptx, 8.35, y, 4.8, 0.95, C.BG3)

        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 8.35, y, w: 0.09, h: 0.95,
          fill: { color: accent }, line: { color: accent },
        })

        slide.addText(c.channel_name, {
          x: 8.58, y: y + 0.06, w: 2.8, h: 0.24,
          fontFace: FONT_B, fontSize: 13, bold: true, color: C.LIGHT, margin: 0,
        })

        slide.addText(safe(c.upload_cadence), {
          x: 8.58, y: y + 0.38, w: 2.5, h: 0.24,
          fontFace: FONT_B, fontSize: 13, color: cadenceColor(c.upload_cadence), margin: 0,
        })

        slide.addText(`~${uploadsPerWeek} videos/week`, {
          x: 10.7, y: y + 0.28, w: 2.1, h: 0.3,
          fontFace: FONT_H, fontSize: 14, bold: true,
          color: C.MUTED, align: "right", margin: 0,
        })
      })

      slide.addText("Upload frequency estimated from recent publishing activity", {
        x: 0.6, y: 5.42, w: 7.3, h: 0.25,
        fontFace: FONT_B, fontSize: 12, color: C.MUTED, align: "center",
      })
    }

    // ================================================
    // SLIDE 7 — ENGAGEMENT ANALYSIS
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Engagement Analysis",
        "Average views, likes, and comments per video across channels"
      )

      // =================================================
      // RAW DATA (NO NORMALIZATION)
      // =================================================

      const companyLabels =
        companies.map(c => c.channel_name)

      const avgViews =
        companies.map(c =>
          Number(c.average_views) || 0
        )

      const avgLikes =
        companies.map(c =>
          Number(c.average_likes) || 0
        )

      const avgComments =
        companies.map(c =>
          Number(c.average_comments) || 0
        )

      // =================================================
      // SMART AXIS SCALING
      // =================================================

      const viewsMax =
        smartAxisMax(avgViews)

      const likesMax =
        smartAxisMax(avgLikes)

      const commentsMax =
        smartAxisMax(avgComments)

      // =================================================
      // CHART POSITIONS
      // =================================================

      const chartW = 4.1
      const chartH = 2.55

      const topY = 1.35
      const bottomY = 4.2

      const gap = 0.22

      const totalWidth =
        (chartW * 3) + (gap * 2)

      const startX =
        (13.33 - totalWidth) / 2

      // =================================================
      // AVG VIEWS CHART
      // =================================================

      slide.addChart(
        pptx.charts.BAR,
        [{
          name: "Avg Views",

          labels: companyLabels,

          values: avgViews,
        }],
        {
          x: startX,
          y: topY,

          w: chartW,
          h: chartH,

          barDir: "col",

          chartColors: [C.ACCENT1],

          chartArea: {
            fill: {
              color: C.BG3
            }
          },

          plotArea: {
            fill: {
              color: C.BG3
            }
          },

          showLegend: false,

          showValue: true,

          dataLabelColor: C.WHITE,
          dataLabelFontSize: 9,

          catAxisLabelColor: C.MUTED,
          valAxisLabelColor: C.MUTED,

          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 8,

          catGridLine: {
            style: "none"
          },

          valGridLine: {
            color: C.BG4,
            style: "solid",
            pt: 0.5
          },

          gapWidthPct: 35,
          valAxisMinVal: 0,
          valAxisMaxVal: viewsMax,
        }
      )

      slide.addText(
        "AVERAGE VIEWS",
        {
          x: startX,
          y: topY - 0.22,

          w: chartW,
          h: 0.2,

          fontFace: FONT_H,
          fontSize: 13,

          bold: true,

          color: C.ACCENT1,

          align: "center",

          margin: 0,
        }
      )

      // =================================================
      // AVG LIKES CHART
      // =================================================

      slide.addChart(
        pptx.charts.BAR,
        [{
          name: "Avg Likes",

          labels: companyLabels,

          values: avgLikes,
        }],
        {
          x: startX + chartW + gap,
          y: topY,

          w: chartW,
          h: chartH,

          barDir: "col",

          chartColors: [C.ACCENT2],

          chartArea: {
            fill: {
              color: C.BG3
            }
          },

          plotArea: {
            fill: {
              color: C.BG3
            }
          },

          showLegend: false,

          showValue: true,

          dataLabelColor: C.WHITE,
          dataLabelFontSize: 9,

          catAxisLabelColor: C.MUTED,
          valAxisLabelColor: C.MUTED,

          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 8,

          catGridLine: {
            style: "none"
          },

          valGridLine: {
            color: C.BG4,
            style: "solid",
            pt: 0.5
          },

          gapWidthPct: 35,
          valAxisMinVal: 0,
          valAxisMaxVal: likesMax,
        }
      )

      slide.addText(
        "AVERAGE LIKES",
        {
          x: startX + chartW + gap,
          y: topY - 0.22,

          w: chartW,
          h: 0.2,

          fontFace: FONT_H,
          fontSize: 13,

          bold: true,

          color: C.ACCENT2,

          align: "center",

          margin: 0,
        }
      )

      // =================================================
      // AVG COMMENTS CHART
      // =================================================

      slide.addChart(
        pptx.charts.BAR,
        [{
          name: "Avg Comments",

          labels: companyLabels,

          values: avgComments,
        }],
        {
          x: startX + (chartW + gap) * 2,
          y: topY,

          w: chartW,
          h: chartH,

          barDir: "col",

          chartColors: [C.ACCENT3],

          chartArea: {
            fill: {
              color: C.BG3
            }
          },

          plotArea: {
            fill: {
              color: C.BG3
            }
          },

          showLegend: false,

          showValue: true,

          dataLabelColor: C.WHITE,
          dataLabelFontSize: 9,

          catAxisLabelColor: C.MUTED,
          valAxisLabelColor: C.MUTED,

          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 8,

          catGridLine: {
            style: "none"
          },

          valGridLine: {
            color: C.BG4,
            style: "solid",
            pt: 0.5
          },

          gapWidthPct: 35,
          valAxisMinVal: 0,
          valAxisMaxVal: commentsMax,
        }
      )

      slide.addText(
        "AVERAGE COMMENTS",
        {
          x: startX + (chartW + gap) * 2,
          y: topY - 0.22,

          w: chartW,
          h: 0.2,

          fontFace: FONT_H,
          fontSize: 13,

          bold: true,

          color: C.ACCENT3,

          align: "center",

          margin: 0,
        }
      )

      // =================================================
      // INSIGHT BOX
      // =================================================

      const insightText =
        safe(data.engagement_analysis)

      card(
        slide,
        pptx,
        0.6,
        bottomY,
        12.1,
        1.45,
        C.BG3
      )

      slide.addText(
        "ENGAGEMENT INSIGHTS",
        {
          x: 0.8,
          y: bottomY + 0.18,

          w: 11.7,
          h: 0.22,

          fontFace: FONT_H,
          fontSize: 15,

          bold: true,

          color: C.ACCENT1,

          align: "center",

          margin: 0,
        }
      )

      slide.addText(
        insightText,
        {
          x: 1.0,
          y: bottomY + 0.55,

          w: 11.2,
          h: 0.6,

          fontFace: FONT_B,
          fontSize: 13.5,

          color: C.LIGHT,

          align: "center",

          valign: "mid",

          margin: 0,
        }
      )

      // =================================================
      // ENGAGEMENT RATE ROW
      // =================================================

      const engRowY = 5.95

      companies.forEach((c, i) => {
        const bw =
          12.5 / companies.length

        const bx =
          0.35 + i * bw

        const accent =
          C.CHART[i % C.CHART.length]

        card(
          slide,
          pptx,
          bx,
          engRowY,
          bw - 0.15,
          0.82,
          C.BG3
        )

        slide.addText(
          `${c.engagement_rate}%`,
          {
            x: bx + 0.08,
            y: engRowY + 0.08,

            w: bw - 0.32,
            h: 0.32,

            fontFace: FONT_H,
            fontSize: 20,

            bold: true,

            color: accent,

            align: "center",

            margin: 0,
          }
        )

        slide.addText(
          c.channel_name,
          {
            x: bx + 0.05,
            y: engRowY + 0.46,

            w: bw - 0.25,
            h: 0.2,

            fontFace: FONT_B,
            fontSize: 11,

            color: C.MUTED,

            align: "center",

            margin: 0,
          }
        )
      })
    }

    // ================================================
    // SLIDE 8 — GAP ANALYSIS
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Gap Analysis",
        "Topics, formats, and audience segments competitors are missing"
      )

      const gapText = safe(data.gap_analysis)
      const gapPoints = gapText
        .split(/(?<=[.!?])\s+/)
        .map(s => s.trim())
        .filter(s => s.length > 10)
        .slice(0, 5)

      const LEFT_X = 0.38
      const LEFT_W = 7.55
      const RIGHT_X = 8.18
      const RIGHT_W = 4.78
      const TOP_Y = 1.22
      const HEADER_H = 0.75
      const PT_H = 0.8
      const PT_GAP = 0.1

      const gapBoxH = HEADER_H + gapPoints.length * (PT_H + PT_GAP) + 0.22

      card(slide, pptx, LEFT_X, TOP_Y, LEFT_W, gapBoxH, C.BG3)

      slide.addShape(pptx.shapes.RECTANGLE, {
        x: LEFT_X, y: TOP_Y, w: LEFT_W, h: 0.1,
        fill: { color: C.ACCENT3 }, line: { color: C.ACCENT3 },
      })

      slide.addText("UNCOVERED OPPORTUNITIES", {
        x: LEFT_X + 0.3, y: TOP_Y + 0.2,
        w: LEFT_W - 0.6, h: 0.34,
        fontFace: FONT_H, fontSize: 17, bold: true, color: C.ACCENT3, align: "left", margin: 0,
      })

      gapPoints.forEach((point, idx) => {
        const rowY = TOP_Y + HEADER_H + idx * (PT_H + PT_GAP)
        const accent = C.CHART[idx % C.CHART.length]

        slide.addShape(pptx.shapes.OVAL, {
          x: LEFT_X + 0.28, y: rowY + 0.24,
          w: 0.16, h: 0.16,
          fill: { color: accent }, line: { color: accent },
        })

        slide.addText(point, {
          x: LEFT_X + 0.56, y: rowY,
          w: LEFT_W - 0.9, h: PT_H,
          fontFace: FONT_B, fontSize: 13.5, color: C.LIGHT,
          valign: "middle", fit: "shrink", margin: 0,
        })
      })

      const chartLabelH = 0.3
      const chartY = 1.08
      const chartH = 2.2

      slide.addText("SUBSCRIBER REACH", {
        x: RIGHT_X, y: chartY, w: RIGHT_W, h: chartLabelH,
        fontFace: FONT_H, fontSize: 14, bold: true,
        color: C.MUTED, align: "center", margin: 0,
      })

      slide.addChart(pptx.charts.BAR,
        [{
          name: "Subscribers",
          labels: companies.map(c => c.channel_name),
          values: companies.map(c => c.subscribers || 0),
        }],
        {
          x: RIGHT_X, y: chartY + chartLabelH,
          w: RIGHT_W, h: chartH,
          barDir: "bar",
          chartColors: companies.map((_, i) => C.CHART[i % C.CHART.length]),
          chartArea: { fill: { color: C.BG3 } },
          plotArea: { fill: { color: C.BG3 } },
          catAxisLabelColor: C.MUTED,
          valAxisLabelColor: C.MUTED,
          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 9,
          catGridLine: { style: "none" },
          valGridLine: { color: "E0D8CC", style: "solid", pt: 0.5 },
          showValue: true,
          dataLabelColor: C.WHITE,
          dataLabelFontSize: 9,
          showLegend: false,
          valAxisMinVal: 0,
        }
      )

      const uploadsStartY = chartY + chartLabelH + chartH + 0.18

      slide.addText("RECENT UPLOADS", {
        x: RIGHT_X, y: uploadsStartY, w: RIGHT_W, h: 0.28,
        fontFace: FONT_H, fontSize: 14, bold: true,
        color: C.MUTED, align: "center", margin: 0,
      })

      const visibleCompanies = companies.slice(0, 3)
      const cardGap = 0.1
      const availH = 7.45 - (uploadsStartY + 0.45)
      const cardH = (availH - ((visibleCompanies.length - 1) * cardGap)) / visibleCompanies.length

      visibleCompanies.forEach((c, i) => {
        const accent = C.CHART[i % C.CHART.length]
        const vids = (c.recent_videos || []).slice(0, 2)
        const cardY = uploadsStartY + 0.38 + i * (cardH + cardGap)

        card(slide, pptx, RIGHT_X, cardY, RIGHT_W, cardH, C.BG3)

        slide.addShape(pptx.shapes.RECTANGLE, {
          x: RIGHT_X, y: cardY, w: 0.08, h: cardH,
          fill: { color: accent }, line: { color: accent },
        })

        slide.addText(c.channel_name, {
          x: RIGHT_X + 0.18, y: cardY + 0.08,
          w: RIGHT_W - 0.28, h: 0.24,
          fontFace: FONT_H, fontSize: 12, bold: true, color: C.ACCENT1, margin: 0,
        })

        vids.forEach((v, vi) => {
          slide.addText(`• ${v.title || "Untitled"}`, {
            x: RIGHT_X + 0.18, y: cardY + 0.34 + vi * 0.24,
            w: RIGHT_W - 0.28, h: 0.22,
            fontFace: FONT_B, fontSize: 9.5, color: C.LIGHT, fit: "shrink", margin: 0,
          })
        })
      })
    }

    // ================================================
    // SLIDE 9 — RECOMMENDATIONS
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Video Marketing Recommendations",
        "Specific, actionable steps to improve YouTube strategy"
      )

      const rawRecs = safe(data.recommendations)
      const recLines = rawRecs
        .split("\n")
        .map(l => l.trim())
        .filter(l => l.length > 5)
        .map(l => l.replace(/^[-•*]\s*/, "").replace(/^\d+\.\s*/, "").trim())
        .filter(l => l.length > 5)
        .slice(0, 6)

      const totalRecs = recLines.length
      const half = Math.ceil(totalRecs / 2)
      const leftRecs = recLines.slice(0, half)
      const rightRecs = recLines.slice(half)

      const CARD_H = 1.05
      const CARD_GAP = 0.22
      const START_Y = 1.38
      const COL_W = 6.1

      const renderRecs = (list, startX, startIndex) => {
        list.forEach((rec, i) => {
          const y = START_Y + i * (CARD_H + CARD_GAP)
          const number = startIndex + i + 1
          const accent = C.CHART[number % C.CHART.length]

          card(slide, pptx, startX, y, COL_W, CARD_H, C.BG3)

          // Left accent bar
          slide.addShape(pptx.shapes.RECTANGLE, {
            x: startX, y, w: 0.09, h: CARD_H,
            fill: { color: accent }, line: { color: accent },
          })

          // Number badge
          slide.addShape(pptx.shapes.OVAL, {
            x: startX + 0.2, y: y + 0.3,
            w: 0.42, h: 0.42,
            fill: { color: accent }, line: { color: accent },
          })

          slide.addText(String(number), {
            x: startX + 0.2, y: y + 0.3,
            w: 0.42, h: 0.42,
            fontFace: FONT_H, fontSize: 13, bold: true,
            color: C.WHITE, align: "center", valign: "middle", margin: 0,
          })

          slide.addText(rec, {
            x: startX + 0.76, y: y + 0.1,
            w: COL_W - 0.92, h: CARD_H - 0.2,
            fontFace: FONT_B, fontSize: 13,
            color: C.LIGHT, valign: "middle", fit: "shrink", margin: 0,
          })
        })
      }

      renderRecs(leftRecs, 0.35, 0)
      renderRecs(rightRecs, 6.82, leftRecs.length)
    }

    // ================================================
    // SLIDE 10 — FINAL RANKING & SCORECARD
    // ================================================
    {
      const slide = setupSlide(
        pptx,
        "Channel Scorecard & Final Ranking",
        "Overall performance scores across key metrics"
      )

      const maxSubs = Math.max(...companies.map(c => c.subscribers || 0))
      const maxViews = Math.max(...companies.map(c => c.average_views || 0))
      const maxEng = Math.max(...companies.map(c => c.engagement_rate || 0))
      const maxFreq = Math.max(...companies.map(c => c.upload_frequency_per_day || 0))

      const scored = companies
        .map(c => {
          const s = maxSubs > 0 ? (c.subscribers / maxSubs) * 25 : 0
          const v = maxViews > 0 ? (c.average_views / maxViews) * 25 : 0
          const e = maxEng > 0 ? (c.engagement_rate / maxEng) * 25 : 0
          const f = maxFreq > 0 ? (c.upload_frequency_per_day / maxFreq) * 25 : 0
          return { ...c, totalScore: Math.round(s + v + e + f) }
        })
        .sort((a, b) => b.totalScore - a.totalScore)

      const tableY = 1.35
      const headers = ["Rank", "Channel", "Subscribers", "Avg Views", "Engagement", "Frequency"]
      const colWidths = [0.9, 3.0, 1.8, 1.8, 1.8, 2.0]

      let cx = 1.0
      card(slide, pptx, 1.0, tableY, 10.8, 0.5, C.ACCENT1, 0.05)

      headers.forEach((h, hi) => {
        slide.addText(h, {
          x: cx + 0.08, y: tableY + 0.12,
          w: colWidths[hi] - 0.12, h: 0.28,
          fontFace: FONT_H, fontSize: 14, bold: true,
          color: C.WHITE,
          align: hi === 0 ? "center" : "left", margin: 0,
        })
        cx += colWidths[hi]
      })

      scored.forEach((c, ri) => {
        const rowY = tableY + 0.6 + ri * 0.78
        const accentColor = C.CHART[ri % C.CHART.length]

        card(slide, pptx, 1.0, rowY, 10.8, 0.7, ri % 2 === 0 ? C.BG3 : C.BG4)

        const medal = ri === 0 ? "🥇" : ri === 1 ? "🥈" : ri === 2 ? "🥉" : `#${ri + 1}`
        const uploadsPerWeek = Math.round((c.upload_frequency_per_day || 0) * 7)
        const vals = [medal, c.channel_name, fmt(c.subscribers), fmt(c.average_views), `${c.engagement_rate}%`, `~${uploadsPerWeek}/week`]

        let vx = 1.0
        vals.forEach((v, vi) => {
          slide.addText(v, {
            x: vx + 0.08, y: rowY + 0.18,
            w: colWidths[vi] - 0.12, h: 0.3,
            fontFace: vi === 1 ? FONT_H : FONT_B,
            fontSize: vi === 1 ? 14 : 13,
            bold: vi === 1,
            color: vi === 1 ? C.ACCENT1 : vi === 0 ? accentColor : C.LIGHT,
            align: vi === 0 ? "center" : "left", margin: 0,
          })
          vx += colWidths[vi]
        })
      })

      const verdict = safe(data.final_ranking).split("\n")[0]
      const verdictLines = Math.ceil(verdict.length / 90)
      const verdictHeight = Math.max(0.7, Math.min(1.2, verdictLines * 0.28))
      const verdictY = tableY + 0.7 + scored.length * 0.78 + 0.35

      card(slide, pptx, 1.0, verdictY, 10.8, verdictHeight, C.BG4)

      // Accent bar on verdict box
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 1.0, y: verdictY, w: 10.8, h: 0.06,
        fill: { color: C.ACCENT1 }, line: { color: C.ACCENT1 },
      })

      slide.addText(verdict, {
        x: 1.25, y: verdictY + 0.18,
        w: 10.2, h: verdictHeight - 0.2,
        fontFace: FONT_B, fontSize: 13, color: C.LIGHT,
        align: "center", valign: "mid",
      })
    }

    // ================================================
    // SAVE
    // ================================================

    const saveDate = new Date()

    const formattedDate =
      `${saveDate.getFullYear()}-${
        String(saveDate.getMonth() + 1).padStart(2, "0")
      }-${
        String(saveDate.getDate()).padStart(2, "0")
      }`

    const fileName =
      `reports/Video Competitor Intelligence Report ${formattedDate}.pptx`

    console.log("WRITING PPT")

    await pptx.writeFile({
      fileName
    })

    console.log("PPT SAVED:", fileName)

    res.json({
      success: true,
      ppt_path: fileName
    })

  } catch (err) {
    console.error(err)
    res.status(500).json({
      success: false,
      error: err.message
    })
  }
})

app.listen(5000, () => console.log("PPT Service running on port 5000"))