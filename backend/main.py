from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv
from datetime import datetime
from openai import OpenAI

import os
import json
import requests
import httpx

app=FastAPI()

load_dotenv()

YOUTUBE_API_KEY=os.getenv("YOUTUBE_API_KEY")
OPENROUTER_API_KEY=os.getenv("OPENROUTER_API_KEY")

client=OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class CompanyRequest(BaseModel):
    company:str
    competitors:list[str]

@app.get("/")
def home():
    return {"message":"Backend is working"}

@app.get("/test-youtube")
def test_youtube():
    if not YOUTUBE_API_KEY:
        return {"error":"YOUTUBE_API_KEY is missing"}

    url="https://www.googleapis.com/youtube/v3/search"
    params={
        "part":"snippet",
        "q":"Nike official",
        "type":"channel",
        "maxResults":1,
        "key":YOUTUBE_API_KEY
    }

    try:
        response=requests.get(url,params=params,timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        return {"error":f"YouTube test failed: {str(e)}"}

cache={}

def safe_int(value,default=0):
    try:
        return int(value)
    except (TypeError,ValueError):
        return default

def parse_iso_datetime(value):
    return datetime.strptime(value,"%Y-%m-%dT%H:%M:%SZ")

def normalize_bullets(value,expected_count):
    if not value or not isinstance(value,str):
        return "\n".join(["- Additional insight unavailable."]*expected_count)

    lines=[line.strip() for line in value.split("\n") if line.strip()]
    bullet_lines=[line if line.startswith("- ") else f"- {line.lstrip('- ').strip()}" for line in lines if line]

    if len(bullet_lines)>=expected_count:
        return "\n".join(bullet_lines[:expected_count])

    text=value.replace("\n"," ").strip()
    sentences=[item.strip() for item in text.split(".") if item.strip()]
    rebuilt=[f"- {sentence}." for sentence in sentences[:expected_count]]

    while len(rebuilt)<expected_count:
        rebuilt.append("- Additional insight unavailable.")

    return "\n".join(rebuilt)

def get_channel_data(company_name):
    if company_name in cache:
        print("USING CACHE:",company_name)
        return cache[company_name]

    if not YOUTUBE_API_KEY:
        print("CHANNEL ERROR: YOUTUBE_API_KEY is missing")
        return None

    try:
        search_url="https://www.googleapis.com/youtube/v3/search"
        search_params={
            "part":"snippet",
            "q":f"{company_name} official",
            "type":"channel",
            "maxResults":1,
            "key":YOUTUBE_API_KEY
        }

        search_response=requests.get(search_url,params=search_params,timeout=30)
        search_response.raise_for_status()
        search_data=search_response.json()

        if "items" not in search_data or not search_data["items"]:
            return None

        channel_id=search_data["items"][0]["id"]["channelId"]

        channel_url="https://www.googleapis.com/youtube/v3/channels"
        channel_params={
            "part":"snippet,statistics,contentDetails",
            "id":channel_id,
            "key":YOUTUBE_API_KEY
        }

        channel_response=requests.get(channel_url,params=channel_params,timeout=30)
        channel_response.raise_for_status()
        channel_data=channel_response.json()

        if "items" not in channel_data or not channel_data["items"]:
            return None

        channel=channel_data["items"][0]
        uploads_playlist_id=channel["contentDetails"]["relatedPlaylists"]["uploads"]

        playlist_url="https://www.googleapis.com/youtube/v3/playlistItems"
        playlist_params={
            "part":"snippet",
            "playlistId":uploads_playlist_id,
            "maxResults":10,
            "key":YOUTUBE_API_KEY
        }

        playlist_response=requests.get(playlist_url,params=playlist_params,timeout=30)
        playlist_response.raise_for_status()
        playlist_data=playlist_response.json()

        if "items" not in playlist_data or not playlist_data["items"]:
            return None

        videos=[]
        for item in playlist_data["items"]:
            try:
                snippet=item["snippet"]
                videos.append({
                    "video_id":snippet["resourceId"]["videoId"],
                    "title":snippet["title"],
                    "published_at":snippet["publishedAt"]
                })
            except Exception:
                continue

        if not videos:
            return None

        videos.sort(key=lambda x:x["published_at"],reverse=True)

        video_ids=[video["video_id"] for video in videos]

        videos_url="https://www.googleapis.com/youtube/v3/videos"
        videos_params={
            "part":"statistics",
            "id":",".join(video_ids),
            "key":YOUTUBE_API_KEY
        }

        videos_response=requests.get(videos_url,params=videos_params,timeout=30)
        videos_response.raise_for_status()
        videos_data=videos_response.json()

        stats_by_id={}
        for item in videos_data.get("items",[]):
            stats_by_id[item["id"]]=item.get("statistics",{})

        for video in videos:
            stats=stats_by_id.get(video["video_id"],{})
            video["views"]=safe_int(stats.get("viewCount"))
            video["likes"]=safe_int(stats.get("likeCount"))
            video["comments"]=safe_int(stats.get("commentCount"))

        total_views=sum(video["views"] for video in videos)
        total_likes=sum(video["likes"] for video in videos)
        total_comments=sum(video["comments"] for video in videos)

        average_views=round(total_views/len(videos),2)
        average_likes=round(total_likes/len(videos),2)
        average_comments=round(total_comments/len(videos),2)

        engagement_rate=0
        if average_views>0:
            engagement_rate=round(((average_likes+average_comments)/average_views)*100,2)

        newest_date=parse_iso_datetime(videos[0]["published_at"])
        oldest_date=parse_iso_datetime(videos[-1]["published_at"])
        days_difference=(newest_date-oldest_date).days

        upload_frequency_per_day=len(videos) if days_difference<=0 else round(len(videos)/days_difference,2)
        uploads_per_week=round(upload_frequency_per_day*7)

        if uploads_per_week>=25:
            cadence="Extremely Active"
        elif uploads_per_week>=10:
            cadence="Very Active"
        elif uploads_per_week>=4:
            cadence="Moderately Active"
        elif uploads_per_week>=1:
            cadence="Occasionally Active"
        else:
            cadence="Low Activity"

        max_views=max((video["views"] for video in videos),default=1)
        scored_videos=[]

        for video in videos:
            views=video["views"]
            likes=video["likes"]
            comments=video["comments"]

            video_engagement=round(((likes+comments)/views)*100,2) if views>0 else 0
            normalized_views=(views/max_views)*100 if max_views>0 else 0
            performance_score=round((0.7*normalized_views)+(0.3*video_engagement),2)

            enriched_video={
                **video,
                "video_engagement":video_engagement,
                "performance_score":performance_score
            }
            scored_videos.append(enriched_video)

        top_video=max(scored_videos,key=lambda x:x["performance_score"])

        result={
            "channel_name":channel["snippet"]["title"],
            "subscribers":safe_int(channel["statistics"].get("subscriberCount")),
            "total_views":safe_int(channel["statistics"].get("viewCount")),
            "total_videos":safe_int(channel["statistics"].get("videoCount")),
            "average_views":average_views,
            "average_likes":average_likes,
            "average_comments":average_comments,
            "engagement_rate":engagement_rate,
            "upload_frequency_per_day":upload_frequency_per_day,
            "uploads_per_week":uploads_per_week,
            "upload_cadence":cadence,
            "top_video":top_video,
            "recent_videos":scored_videos
        }

        cache[company_name]=result
        return result

    except requests.RequestException as e:
        print("CHANNEL REQUEST ERROR:",str(e))
        return None
    except Exception as e:
        print("CHANNEL ERROR:",str(e))
        return None

def generate_full_report(results):
    compact_results=[]

    for company in results:
        compact_results.append({
            "company":company["channel_name"],
            "subscribers":company["subscribers"],
            "total_videos":company["total_videos"],
            "avg_views":company["average_views"],
            "avg_likes":company["average_likes"],
            "avg_comments":company["average_comments"],
            "engagement":company["engagement_rate"],
            "cadence":company["upload_cadence"],
            "uploads_per_week":company["uploads_per_week"],
            "top_video":company["top_video"]["title"],
            "recent_videos":[video["title"] for video in company.get("recent_videos",[])]
        })

    prompt=f"""
You are a senior YouTube marketing strategist preparing a boardroom-quality competitor analysis report.

Analyze the following data for {len(compact_results)} YouTube channels and return STRICT VALID JSON ONLY.
Do NOT include any text before or after the JSON block.
Do NOT wrap the JSON in markdown code fences.

IMPORTANT FORMATTING RULES:
- "missing_opportunities" must contain EXACTLY 5 lines, each line starting with "- " (hyphen space), one per line, separated by \\n.
- "recommendations" must contain EXACTLY 6 lines, each line starting with "- " (hyphen space), one per line, separated by \\n.
- All other fields must be plain paragraph text with no bullet points or line breaks.

Return this exact JSON structure:
{{
  "executive_summary":"3-5 sentence plain paragraph summary.",
  "engagement_analysis":"2-3 sentence plain paragraph engagement analysis.",
  "content_themes":"2-4 sentence plain paragraph content analysis.",
  "gap_analysis":"5-7 sentence plain paragraph market gap analysis.",
  "missing_opportunities":"- Line one opportunity here.\\n- Line two opportunity here.\\n- Line three opportunity here.\\n- Line four opportunity here.\\n- Line five opportunity here.",
  "recommendations":"- Recommendation one here.\\n- Recommendation two here.\\n- Recommendation three here.\\n- Recommendation four here.\\n- Recommendation five here.\\n- Recommendation six here.",
  "final_ranking":"Rank channels from best to worst with brief reason for each."
}}

DATA:
{json.dumps(compact_results,indent=2)}
"""

    fallback_report={
        "executive_summary":"AI analysis unavailable.",
        "engagement_analysis":"AI analysis unavailable.",
        "content_themes":"AI analysis unavailable.",
        "gap_analysis":"AI analysis unavailable.",
        "missing_opportunities":
            "AI analysis unavailable."
        ,
        "recommendations":
            "Ai analysis unavailable.",
        "final_ranking":"AI analysis unavailable."
    }

    try:
        if not OPENROUTER_API_KEY:
            raise ValueError("OPENROUTER_API_KEY is missing")

        completion=client.chat.completions.create(
            model="google/gemma-4-31b-it",
            messages=[
                {
                    "role":"user",
                    "content":prompt
                }
            ],
            temperature=0.4,
            max_tokens=2000,
            response_format={
                "type":"json_schema",
                "json_schema":{
                    "name":"youtube_competitor_report",
                    "strict":True,
                    "schema":{
                        "type":"object",
                        "properties":{
                            "executive_summary":{"type":"string"},
                            "engagement_analysis":{"type":"string"},
                            "content_themes":{"type":"string"},
                            "gap_analysis":{"type":"string"},
                            "missing_opportunities":{"type":"string"},
                            "recommendations":{"type":"string"},
                            "final_ranking":{"type":"string"}
                        },
                        "required":[
                            "executive_summary",
                            "engagement_analysis",
                            "content_themes",
                            "gap_analysis",
                            "missing_opportunities",
                            "recommendations",
                            "final_ranking"
                        ],
                        "additionalProperties":False
                    }
                }
            }
        )

        response=completion.choices[0].message.content.strip()
        response=response.replace("```json","").replace("```","").strip()

        json_start=response.find("{")
        json_end=response.rfind("}")+1

        if json_start==-1 or json_end==0:
            raise ValueError("No JSON object found in model response")

        response=response[json_start:json_end]
        report_json=json.loads(response)

        report_json["missing_opportunities"]=normalize_bullets(
            report_json.get("missing_opportunities",""),
            5
        )
        report_json["recommendations"]=normalize_bullets(
            report_json.get("recommendations",""),
            6
        )

        for field in [
            "executive_summary",
            "engagement_analysis",
            "content_themes",
            "gap_analysis",
            "final_ranking"
        ]:
            value=report_json.get(field,"")
            if not isinstance(value,str) or not value.strip():
                report_json[field]=fallback_report[field]
            else:
                report_json[field]=value.replace("\n"," ").strip()

        return report_json

    except Exception as e:
        print("AI ERROR:",str(e))
        return fallback_report

@app.post("/analyze")
def analyze(data:CompanyRequest):
    companies=[data.company]+data.competitors
    results=[]
    failed_companies=[]

    for company in companies:
        company=company.strip()
        if not company:
            continue

        try:
            company_data=get_channel_data(company)
            if company_data:
                results.append(company_data)
            else:
                failed_companies.append(company)
        except Exception as e:
            print(f"ERROR {company}:",str(e))
            failed_companies.append(company)

    if not results:
        return {
            "error":"No valid YouTube channels found.",
            "failed_companies":failed_companies
        }

    leader_by_subscribers=max(results,key=lambda x:x["subscribers"])
    leader_by_engagement=max(results,key=lambda x:x["engagement_rate"])
    leader_by_average_views=max(results,key=lambda x:x["average_views"])
    most_active_company=max(results,key=lambda x:x["uploads_per_week"])

    ai_sections=generate_full_report(results)

    report_response={
        "executive_summary":ai_sections["executive_summary"],
        "engagement_analysis":ai_sections["engagement_analysis"],
        "content_themes":ai_sections["content_themes"],
        "gap_analysis":ai_sections["gap_analysis"],
        "missing_opportunities":ai_sections["missing_opportunities"],
        "recommendations":ai_sections["recommendations"],
        "final_ranking":ai_sections["final_ranking"],
        "leaderboards":{
            "highest_subscribers":{
                "company":leader_by_subscribers["channel_name"]
            },
            "highest_engagement":{
                "company":leader_by_engagement["channel_name"]
            },
            "highest_average_views":{
                "company":leader_by_average_views["channel_name"]
            },
            "most_active_channel":{
                "company":most_active_company["channel_name"]
            }
        },
        "companies_analyzed":results,
        "failed_companies":failed_companies
    }

    try:
        ppt_response=httpx.post(
            "http://127.0.0.1:5000/generate-ppt",
            json=report_response,
            timeout=120
        )

        try:
            ppt_data=ppt_response.json()
        except Exception as json_err:
            raise RuntimeError(
                f"Invalid PPT service response: {json_err} | body: {ppt_response.text}"
            ) from json_err

        if ppt_response.status_code!=200:
            raise RuntimeError(
                f"PPT service returned {ppt_response.status_code}: {ppt_data}"
            )

        ppt_path=ppt_data.get("ppt_path")
        if not ppt_path:
            raise KeyError(f"Missing ppt_path in PPT service response: {ppt_data}")

        report_response["ppt_path"]=ppt_path

    except Exception as e:
        print("PPT SERVICE ERROR:",str(e))
        report_response["ppt_error"]=str(e)

    return report_response