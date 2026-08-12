repo: codingsfst-cyber/smart.flex
branch: main
path: index.html

## Last sync
date: 2026-08-12T12:36:35Z

### Updated in this project
- Working index.html rewritten on the repo's palette; bottom-nav icons and the SF logo mark lifted verbatim from index.html
- Thai voice coaching, per-joint green/red skeleton, dashed target-posture guide, session timer, stop button
- Built SmartFlex mobile prototype (Thai + English) on the repo's palette, radii and Noto Sans Thai type
- Full flow: home → camera setup → live session with nudges → summary + stretch sheet → session report
- Two navigation models (card stack / session-first) switchable in-design
- Added ตรวจเดี่ยว / ตรวจหมู่ modes, baseline calibration, and Google Sheets sync + offline queue screens
- Auto angle detection (side / front / full body) with 3-second re-framing overlay

## Screen map
| Project screen | Repo source |
| --- | --- |
| SmartFlex.dc.html — all screens | index.html (`:root` tokens, `.card`, `.bottom-nav`, `.score-ring`, `.metric-box`, `.issue-pill`) |
| index.html (working app) | index.html (`:root`, `.bottom-nav` SVG paths lines 769–786, favicon line 8, sheet columns line 2482) |
| Code.gs (Apps Script backend) | index.html (`APPS_SCRIPT_URL`, 15-column header) |
| Nav A / Nav B home explorations | index.html (`.top-nav`, `.bottom-nav`, `.kpi-box`) |
