import { useState } from "react"

function App() {
  const [company,setCompany]=useState("")
  const [competitor1,setCompetitor1]=useState("")
  const [competitor2,setCompetitor2]=useState("")
  const [competitor3,setCompetitor3]=useState("")
  const [competitor4,setCompetitor4]=useState("")
  const [reportData,setReportData]=useState(null)
  const [loading,setLoading]=useState(false)
  const [error,setError]=useState("")

  function safeText(value) {
    if(value===null||value===undefined||value==="") return "No data available"
    if(typeof value==="string") return value
    return JSON.stringify(value)
  }

  function convertToPoints(text) {
    if(!text||typeof text!=="string") return []
    return text
      .split(/\n|•|(?<=\.)\s+(?=[A-Z-])/)
      .map(item=>item.replace(/^[-•\d.\s]+/,"").trim())
      .filter(item=>item.length>5)
  }

  const generateReport = async () => {
    try {
      setLoading(true)

      // =========================================
      // CREATE COMPETITOR ARRAY
      // =========================================

      const competitorList = [competitor1, competitor2, competitor3, competitor4]
        .filter(c => c.trim() !== "")

      // =========================================
      // STEP 1 — CALL FASTAPI ANALYZE
      // =========================================

      const analysisResponse = await fetch(
        "https://video-competitor-intelligence-i31j.onrender.com/analyze",
        {
          method: "POST",

          headers: {
            "Content-Type": "application/json",
          },

          body: JSON.stringify({
            company: company.trim(),
            competitors: competitorList,
          }),
        }
      )

      if (!analysisResponse.ok) {
        throw new Error(`Backend error: ${analysisResponse.status} ${analysisResponse.statusText}`)
      }

      const analysisData =
        await analysisResponse.json()

      console.log("ANALYSIS DATA:")
      console.log(analysisData)

      // =========================================
      // VERY IMPORTANT CHECK
      // =========================================

      if (!analysisData.companies_analyzed) {
        throw new Error("Backend did not return companies_analyzed")
      }

      // =========================================
      // STEP 2 — SEND TO PPT SERVICE
      // =========================================

      const pptResponse = await fetch(
        "https://ppt-service.onrender.com/generate-ppt",
        {
          method: "POST",

          headers: {
            "Content-Type": "application/json",
          },

          body: JSON.stringify(analysisData),
        }
      )

      if (!pptResponse.ok) {
        throw new Error(`PPT service error: ${pptResponse.status} ${pptResponse.statusText}`)
      }

      const pptData =
        await pptResponse.json()

      console.log("PPT DATA:")
      console.log(pptData)

      if (pptData.error) {
        throw new Error(pptData.error)
      }

      setReportData({
        analysis: analysisData,
        ppt_path: pptData.ppt_path
      })

    } catch (err) {
      console.error("Full error:", err)
      const errorMsg = err?.message || "Unknown error"
      setError(`Failed: ${errorMsg}`)
      alert(`Failed to generate report: ${errorMsg}`)
    }

    finally {
      setLoading(false)
    }
  }

  return (
    <div style={pageStyle}>
      <div style={heroSection}>
        <div style={overlay}></div>
        <div style={heroContent}>
          <img src="/youtube.jpeg" alt="YouTube" style={youtubeLogo} />
          <h1 style={mainTitle}>Video Competitor Intelligence and Report Generator</h1>
          <p style={subTitle}>AI-Powered YouTube Competitor Analysis</p>
        </div>
      </div>

      <div style={formContainer}>
        <div style={topInputContainer}>
          <input
            type="text"
            placeholder="Your Company"
            value={company}
            onChange={e=>setCompany(e.target.value)}
            style={topInputStyle}
          />
        </div>

        <div style={competitorGrid}>
          <input
            type="text"
            placeholder="Competitor 1"
            value={competitor1}
            onChange={e=>setCompetitor1(e.target.value)}
            style={inputStyle}
          />
          <input
            type="text"
            placeholder="Competitor 2"
            value={competitor2}
            onChange={e=>setCompetitor2(e.target.value)}
            style={inputStyle}
          />
          <input
            type="text"
            placeholder="Competitor 3"
            value={competitor3}
            onChange={e=>setCompetitor3(e.target.value)}
            style={inputStyle}
          />
          <input
            type="text"
            placeholder="Competitor 4"
            value={competitor4}
            onChange={e=>setCompetitor4(e.target.value)}
            style={inputStyle}
          />
        </div>

        <div style={{textAlign:"center"}}>
          <button onClick={generateReport} style={buttonStyle} disabled={loading}>
            {loading?"Generating...":"Generate Report"}
          </button>
        </div>

        {error&&<div style={errorStyle}>{error}</div>}
      </div>

      {loading&&(
        <div style={loadingStyle}>
          Analyzing competitors and generating report...
        </div>
      )}

      {reportData&&(
        <div style={reportContainer}>
          <div style={cardStyle}>
            <h2 style={headingStyle}>Executive Summary</h2>
            <p style={textStyle}>{safeText(reportData.analysis?.executive_summary)}</p>
          </div>

          <div style={kpiContainer}>
            <div style={kpiStyle}>
              <h3 style={kpiHeading}>Highest Subscribers</h3>
              <p style={kpiValue}>
                {reportData?.analysis?.leaderboards?.highest_subscribers?.company||"N/A"}
              </p>
            </div>

            <div style={kpiStyle}>
              <h3 style={kpiHeading}>Highest Engagement</h3>
              <p style={kpiValue}>
                {reportData?.analysis?.leaderboards?.highest_engagement?.company||"N/A"}
              </p>
            </div>

            <div style={kpiStyle}>
              <h3 style={kpiHeading}>Highest Avg Views</h3>
              <p style={kpiValue}>
                {reportData?.analysis?.leaderboards?.highest_average_views?.company||"N/A"}
              </p>
            </div>

            <div style={kpiStyle}>
              <h3 style={kpiHeading}>Most Active</h3>
              <p style={kpiValue}>
                {reportData?.analysis?.leaderboards?.most_active_channel?.company||"N/A"}
              </p>
            </div>
          </div>

          <div style={companyGrid}>
            {reportData?.analysis?.companies_analyzed?.map((companyItem,index)=>(
              <div key={index} style={companyCard}>
                <h2 style={companyTitle}>{companyItem.channel_name}</h2>

                <div style={metricRow}>
                  <span style={metricLabel}>Subscribers</span>
                  <span style={metricValue}>
                    {companyItem.subscribers?.toLocaleString()||"N/A"}
                  </span>
                </div>

                <div style={metricRow}>
                  <span style={metricLabel}>Avg Views</span>
                  <span style={metricValue}>
                    {companyItem.average_views?.toLocaleString()||"N/A"}
                  </span>
                </div>

                <div style={metricRow}>
                  <span style={metricLabel}>Engagement</span>
                  <span style={metricValue}>
                    {companyItem.engagement_rate!==undefined?`${companyItem.engagement_rate}%`:"N/A"}
                  </span>
                </div>

                <div style={metricRow}>
                  <span style={metricLabel}>Upload Cadence</span>
                  <span style={metricValue}>
                    {companyItem.upload_cadence||"N/A"}
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div style={cardStyle}>
            <h2 style={headingStyle}>Content Themes</h2>
            <ul style={bulletListStyle}>
              {convertToPoints(reportData.analysis?.content_themes).map((point,index)=>(
                <li key={index} style={bulletItem}>{point}</li>
              ))}
            </ul>
          </div>

          <div style={cardStyle}>
            <h2 style={headingStyle}>Gap Analysis</h2>
            <ul style={bulletListStyle}>
              {convertToPoints(reportData.analysis?.gap_analysis).map((point,index)=>(
                <li key={index} style={bulletItem}>{point}</li>
              ))}
            </ul>
          </div>

          <div style={cardStyle}>
            <h2 style={headingStyle}>Recommendations</h2>
            <ul style={bulletListStyle}>
              {convertToPoints(reportData.analysis?.recommendations).map((point,index)=>(
                <li key={index} style={bulletItem}>{point}</li>
              ))}
            </ul>
          </div>

          {reportData.ppt_path&&(
            <div style={{textAlign:"center"}}>
              <a href={'https://ppt-service.onrender.com' + reportData.ppt_path} download>
                <button style={downloadButton}>Download PowerPoint Report</button>
              </a>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

const pageStyle={
  background:"linear-gradient(180deg,#F6F1E7 0%,#EEE4D2 100%)",
  minHeight:"100vh",
  color:"#2A1E17",
  fontFamily:"'Poppins','Segoe UI',sans-serif"
}

const heroSection={
  height:"52vh",
  position:"relative",
  display:"flex",
  justifyContent:"center",
  alignItems:"center",
  overflow:"hidden",
  background:"linear-gradient(rgba(255,248,240,0.80),rgba(255,248,240,0.88)),url('/youtube.jpeg')",
  backgroundSize:"cover",
  backgroundPosition:"center"
}

const overlay={
  position:"absolute",
  inset:0,
  backdropFilter:"blur(1px)"
}

const heroContent={
  position:"relative",
  zIndex:2,
  textAlign:"center"
}

const youtubeLogo={
  width:"140px",
  marginBottom:"24px",
  filter:"drop-shadow(0px 12px 30px rgba(0,0,0,0.25))"
}

const mainTitle={
  fontSize:"72px",
  fontWeight:"900",
  marginBottom:"12px",
  color:"#661111",
  letterSpacing:"-2px"
}

const subTitle={
  fontSize:"22px",
  color:"#7A3D1B",
  fontWeight:"500"
}

const formContainer={
  maxWidth:"1100px",
  margin:"-70px auto 0 auto",
  position:"relative",
  zIndex:5,
  background:"rgba(255,255,255,0.72)",
  backdropFilter:"blur(14px)",
  padding:"40px",
  borderRadius:"28px",
  boxShadow:"0 20px 60px rgba(0,0,0,0.18)"
}

const topInputContainer={
  display:"flex",
  justifyContent:"center",
  marginBottom:"28px"
}

const topInputStyle={
  width:"65%",
  padding:"20px",
  borderRadius:"18px",
  border:"1px solid rgba(0,0,0,0.08)",
  backgroundColor:"rgba(255,255,255,0.88)",
  fontSize:"18px",
  color:"#2A1E17",
  outline:"none",
  boxShadow:"0 6px 18px rgba(0,0,0,0.08)"
}

const competitorGrid={
  display:"grid",
  gridTemplateColumns:"1fr 1fr",
  gap:"22px",
  marginBottom:"36px"
}

const inputStyle={
  padding:"18px",
  borderRadius:"18px",
  border:"1px solid rgba(0,0,0,0.08)",
  backgroundColor:"rgba(255,255,255,0.88)",
  fontSize:"17px",
  color:"#2A1E17",
  outline:"none",
  boxShadow:"0 6px 18px rgba(0,0,0,0.08)"
}

const buttonStyle={
  padding:"18px 42px",
  borderRadius:"18px",
  border:"none",
  background:"linear-gradient(135deg,#8B0000,#C21807)",
  color:"#FFFFFF",
  fontWeight:"700",
  fontSize:"18px",
  cursor:"pointer",
  boxShadow:"0 12px 30px rgba(139,0,0,0.30)",
  transition:"0.3s ease"
}

const errorStyle={
  marginTop:"20px",
  textAlign:"center",
  color:"#B00020",
  fontSize:"16px",
  fontWeight:"600"
}

const loadingStyle={
  marginTop:"30px",
  textAlign:"center",
  color:"#7A3D1B",
  fontSize:"18px"
}

const reportContainer={
  marginTop:"60px",
  maxWidth:"1400px",
  marginLeft:"auto",
  marginRight:"auto",
  padding:"0 30px 60px 30px"
}

const cardStyle={
  background:"rgba(255,255,255,0.80)",
  padding:"36px",
  borderRadius:"24px",
  marginTop:"30px",
  boxShadow:"0 12px 34px rgba(0,0,0,0.12)"
}

const headingStyle={
  color:"#8B0000",
  marginBottom:"22px",
  fontSize:"30px",
  fontWeight:"800"
}

const textStyle={
  lineHeight:"2",
  fontSize:"18px",
  color:"#2A1E17"
}

const kpiContainer={
  display:"grid",
  gridTemplateColumns:"repeat(auto-fit,minmax(230px,1fr))",
  gap:"22px",
  marginTop:"35px"
}

const kpiStyle={
  background:"rgba(255,255,255,0.82)",
  padding:"30px",
  borderRadius:"22px",
  textAlign:"center",
  boxShadow:"0 10px 28px rgba(0,0,0,0.10)"
}

const kpiHeading={
  color:"#7A3D1B",
  marginBottom:"16px",
  fontSize:"18px"
}

const kpiValue={
  fontSize:"28px",
  fontWeight:"800",
  color:"#8B0000"
}

const companyGrid={
  marginTop:"40px",
  display:"grid",
  gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",
  gap:"24px"
}

const companyCard={
  background:"rgba(255,255,255,0.82)",
  padding:"30px",
  borderRadius:"22px",
  boxShadow:"0 10px 28px rgba(0,0,0,0.10)"
}

const companyTitle={
  fontSize:"28px",
  marginBottom:"24px",
  color:"#8B0000"
}

const metricRow={
  display:"flex",
  justifyContent:"space-between",
  marginBottom:"16px",
  paddingBottom:"12px",
  borderBottom:"1px solid rgba(0,0,0,0.08)"
}

const metricLabel={
  color:"#7A3D1B",
  fontSize:"15px"
}

const metricValue={
  color:"#2A1E17",
  fontWeight:"700"
}

const bulletListStyle={
  paddingLeft:"26px",
  lineHeight:"2.1",
  fontSize:"17px",
  color:"#2A1E17"
}

const bulletItem={
  marginBottom:"14px"
}

const downloadButton={
  marginTop:"45px",
  padding:"18px 40px",
  borderRadius:"18px",
  border:"none",
  background:"linear-gradient(135deg,#8B0000,#C21807)",
  color:"#FFFFFF",
  fontWeight:"700",
  fontSize:"17px",
  cursor:"pointer",
  boxShadow:"0 12px 28px rgba(139,0,0,0.25)"
}

export default App