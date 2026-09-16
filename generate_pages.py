"""2026 한신대학교 채용박람회 — 동문 멘토별 사전질문 페이지 생성기.

구글 폼 응답 시트의 열 제목에서 '기업_직무_학과' 3단 구조인 열만 동문 멘토로 보고
c/<번호>.html 과 c/index.html(목록)을 만든다.
학생 응답은 페이지에서 실시간으로 시트를 읽으므로 재생성이 필요 없다.
멘토가 추가·변경됐을 때만 다시 실행:

    python generate_pages.py
"""
import csv, html, io, os, re, urllib.request

SHEET_ID = "17rF8qsEcRVtfN8nGmi7P3-7-xABT2fYbAH1xxfQ3iAI"
GID = "2037502648"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid={GID}"
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "c")
BASE = "https://polanyi082.github.io/HANSHIN-UNIVERSITY_JOB-MENTOR/c/"

EVENT = "2026 한신대학교 채용박람회 · 대학일자리플러스센터"


def fetch():
    """[(label, name, job, major, 신청인원), ...] — 시트 열 순서 유지, 중복 열 병합."""
    with urllib.request.urlopen(CSV_URL) as r:
        table = list(csv.reader(io.StringIO(r.read().decode("utf-8"))))
    head, body = table[0], table[1:]

    cols = {}                       # label -> [열 index]
    for i, h in enumerate(head):
        m = re.search(r"\[([^\]]*)\]\s*$", h)
        if not m:
            continue
        label = m.group(1).strip()
        parts = [p.strip() for p in label.split("_")]
        if len(parts) != 3 or not all(parts):
            continue                # 현장면접·일반 채용컨설팅 열은 제외
        cols.setdefault(label, []).append(i)

    out = []
    for label, idxs in cols.items():
        n = sum(1 for r in body if any(i < len(r) and r[i].strip() for i in idxs))
        name, job, major = [p.strip() for p in label.split("_")]
        out.append((label, name, job, major, n))
    return out


PAGE = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="mentor.css">
</head>
<body data-mentor="{label}">
<header>
  <div class="hd">
    <span class="badge">2026 한신대학교 채용박람회 · 동문 멘토링</span><span class="booth">멘토 {num}</span>
    <h1>{name}<small>{job} · 동문 {major}</small></h1>
    <div class="sub">{event} · <a href="index.html">전체 멘토 목록</a></div>
  </div>
</header>

<div class="wrap">
  <div class="solo">
    <div class="stats">
      <div class="stat"><b id="s1">–</b><span>신청 학생</span></div>
      <div class="stat o"><b id="s2">–</b><span>사전질문 1</span></div>
      <div class="stat y"><b id="s3">–</b><span>사전질문 2</span></div>
      <div class="stat g"><b id="s5">–</b><span>사전질문 3</span></div>
      <div class="stat w"><b id="s4" style="font-size:14px;padding-top:5px">–</b><span>최종 동기화</span></div>
      <button id="reload">↻ 새로고침</button>
    </div>
  </div>

  <h2>사전질문 리스트</h2>
  <div class="note">전공(계열) 오름차순 정렬 · 구글 폼 응답 시트 실시간 연동</div>
  <div class="tablebox">
    <div class="scroll"><table>
      <thead><tr>
        <th class="idx">#</th><th>전공(계열)</th><th>이름</th><th>학번</th><th>희망 시간대</th>
        <th style="width:26%">사전질문 1 — 직무 역량 · 인재상</th>
        <th style="width:26%">사전질문 2 — 서류/면접 · 어학 · 자격증</th>
        <th style="width:22%">사전질문 3 — 기타 궁금한 점</th>
      </tr></thead>
      <tbody id="tb"><tr><td colspan="8" class="empty">데이터를 불러오는 중입니다…</td></tr></tbody>
    </table></div>
  </div>

  <h2>핵심 키워드 워드클라우드</h2>
  <div class="note">이 멘토를 신청한 학생들의 질문에서 추출한 키워드 · 글자 크기 = 언급 빈도</div>
  <div class="clouds">
    <div class="cloudcard">
      <h3>사전질문 1 키워드</h3>
      <div class="cap">희망 직무의 핵심 실무 역량 · 인재상 · 근무환경</div>
      <div class="circle" id="c1"></div>
      <div class="chips" id="k1"></div>
    </div>
    <div class="cloudcard q2">
      <h3>사전질문 2 키워드</h3>
      <div class="cap">서류/면접 준비 전략 · 어학 · 자격증 · 포트폴리오</div>
      <div class="circle q2" id="c2"></div>
      <div class="chips" id="k2"></div>
    </div>
    <div class="cloudcard q3">
      <h3>사전질문 3 키워드</h3>
      <div class="cap">그 외 다양한 궁금한 점</div>
      <div class="circle q3" id="c3"></div>
      <div class="chips" id="k3"></div>
    </div>
  </div>

  <footer>
    데이터 출처: 구글 폼 응답 시트 「참여학생 사전신청 설문지(응답)」 · 5분마다 자동 갱신<br>
    <a href="index.html">← 전체 멘토 목록</a> · <a href="../">통합 보드</a>
  </footer>
</div>
<script src="mentor.js"></script>
</body>
</html>
"""

HUB = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>2026 한신대학교 채용박람회 · 동문 멘토별 사전질문</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="mentor.css">
</head>
<body>
<header>
  <div class="hd">
    <span class="badge">2026 한신대학교 채용박람회 · 대학일자리플러스센터</span>
    <h1>동문 멘토별 <b>사전질문</b> 페이지<small>멘토님께 각 링크를 개별 안내해 주세요</small></h1>
    <div class="sub">{event}</div>
  </div>
</header>
<div class="wrap">
  <div class="solo">
    <div class="stats">
      <div class="stat"><b id="t1">–</b><span>동문 멘토</span></div>
      <div class="stat o"><b id="t2">–</b><span>멘토링 신청 학생</span></div>
      <div class="stat y"><b id="t3">–</b><span>전체 설문 응답</span></div>
      <div class="stat w"><b id="s4" style="font-size:14px;padding-top:5px">–</b><span>최종 동기화</span></div>
      <button id="reload">↻ 새로고침</button>
    </div>
  </div>
  <div id="hub"><div class="empty">멘토 목록을 불러오는 중입니다…</div></div>
  <footer>
    데이터 출처: 구글 폼 응답 시트 「참여학생 사전신청 설문지(응답)」 · 5분마다 자동 갱신<br>
    <a href="../">통합 보드(멘토 전체 한 화면)</a>
  </footer>
</div>
<script src="mentor.js"></script>
</body>
</html>
"""

BOARD = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>2026 한신대학교 채용박람회 · 동문 멘토 사전질문 보드</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="c/mentor.css">
</head>
<body>
<header>
  <div class="hd">
    <span class="badge">2026 한신대학교 채용박람회 · 대학일자리플러스센터</span>
    <h1>동문 멘토 <b>사전질문</b> 보드</h1>
    <div class="sub">{event} · <a href="c/">멘토별 개별 페이지 목록</a></div>
  </div>
</header>

<div class="wrap">
  <div class="panel">
    <div class="ctrl">
      <div class="field">
        <label for="q">멘토 검색</label>
        <input id="q" type="search" placeholder="기업명 · 직무 · 학과 (예: 카카오, 사회복지)" autocomplete="off">
      </div>
      <div class="field" style="flex:2">
        <label for="sel">동문 멘토 (기업_직무_동문 학과)</label>
        <select id="sel"><option>불러오는 중…</option></select>
      </div>
    </div>
    <div class="stats">
      <div class="stat"><b id="s1">–</b><span>신청 학생</span></div>
      <div class="stat o"><b id="s2">–</b><span>사전질문 1</span></div>
      <div class="stat y"><b id="s3">–</b><span>사전질문 2</span></div>
      <div class="stat g"><b id="s5">–</b><span>사전질문 3</span></div>
      <div class="stat w"><b id="s4" style="font-size:14px;padding-top:5px">–</b><span>최종 동기화</span></div>
      <button id="reload">↻ 새로고침</button>
    </div>
  </div>

  <h2>사전질문 리스트</h2>
  <div class="note">전공(계열) 오름차순 정렬 · 구글 폼 응답 시트 실시간 연동</div>
  <div class="tablebox">
    <div class="scroll"><table>
      <thead><tr>
        <th class="idx">#</th><th>전공(계열)</th><th>이름</th><th>학번</th><th>희망 시간대</th>
        <th style="width:26%">사전질문 1 — 직무 역량 · 인재상</th>
        <th style="width:26%">사전질문 2 — 서류/면접 · 어학 · 자격증</th>
        <th style="width:22%">사전질문 3 — 기타 궁금한 점</th>
      </tr></thead>
      <tbody id="tb"><tr><td colspan="8" class="empty">데이터를 불러오는 중입니다…</td></tr></tbody>
    </table></div>
  </div>

  <h2>핵심 키워드 워드클라우드</h2>
  <div class="note">선택한 멘토를 신청한 학생들의 질문에서 추출한 키워드 · 글자 크기 = 언급 빈도</div>
  <div class="clouds">
    <div class="cloudcard">
      <h3>사전질문 1 키워드</h3>
      <div class="cap">희망 직무의 핵심 실무 역량 · 인재상 · 근무환경</div>
      <div class="circle" id="c1"></div>
      <div class="chips" id="k1"></div>
    </div>
    <div class="cloudcard q2">
      <h3>사전질문 2 키워드</h3>
      <div class="cap">서류/면접 준비 전략 · 어학 · 자격증 · 포트폴리오</div>
      <div class="circle q2" id="c2"></div>
      <div class="chips" id="k2"></div>
    </div>
    <div class="cloudcard q3">
      <h3>사전질문 3 키워드</h3>
      <div class="cap">그 외 다양한 궁금한 점</div>
      <div class="circle q3" id="c3"></div>
      <div class="chips" id="k3"></div>
    </div>
  </div>

  <footer>
    데이터 출처: 구글 폼 응답 시트 「참여학생 사전신청 설문지(응답)」 · 5분마다 자동 갱신<br>
    설문 응답이 추가되면 자동 반영됩니다.
  </footer>
</div>
<script src="c/mentor.js"></script>
</body>
</html>
"""


def main():
    os.makedirs(OUT, exist_ok=True)
    mentors = fetch()
    for i, (label, name, job, major, n) in enumerate(mentors, 1):
        num = f"{i:02d}"
        page = PAGE.format(
            title=f"{name} · 동문 멘토 사전질문 · 2026 한신대학교 채용박람회",
            label=html.escape(label, quote=True), num=num,
            name=html.escape(name), job=html.escape(job), major=html.escape(major),
            event=EVENT,
        )
        with io.open(os.path.join(OUT, f"{num}.html"), "w", encoding="utf-8", newline="\n") as f:
            f.write(page)
    with io.open(os.path.join(OUT, "index.html"), "w", encoding="utf-8", newline="\n") as f:
        f.write(HUB.format(event=EVENT))
    with io.open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8", newline="\n") as f:
        f.write(BOARD.format(event=EVENT))

    lines = ["번호\t기업명\t직무\t동문 학과\t신청\tURL"]
    for i, (label, name, job, major, n) in enumerate(mentors, 1):
        lines.append(f"{i:02d}\t{name}\t{job}\t{major}\t{n}\t{BASE}{i:02d}.html")
    with io.open(os.path.join(ROOT, "URL목록.tsv"), "w", encoding="utf-8-sig", newline="\n") as f:
        f.write("\n".join(lines) + "\n")
    print(f"{len(mentors)}명 멘토 페이지 생성 완료 -> {OUT}")


if __name__ == "__main__":
    main()
