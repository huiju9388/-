"""
GitHub Contents API 배포 헬퍼.
사용법:
  from deploy import put_file
  put_file("data/dashboard-data.json", "output/dashboard-data.json", "주간 업데이트: dashboard-data")

주의: SHA는 매번 새로 조회한다 (직전 단계 SHA 재사용 금지 — 중간 커밋으로 무효화될 수 있음).
배포 검증은 반드시 Contents API로 재조회해서 확인 (web_fetch/raw URL은 CDN 캐싱으로 stale 가능).
"""
import base64, json, urllib.request

REPO = "huiju9388/-"


def _req(method, url, token, data=None):
    r = urllib.request.Request(url, data=data, method=method)
    r.add_header("Authorization", f"token {token}")
    r.add_header("Accept", "application/vnd.github.v3+json")
    r.add_header("User-Agent", "dashboard-updater")
    if data:
        r.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(r) as resp:
        return json.load(resp)


def put_file(path, local_path, msg, token):
    meta = _req("GET", f"https://api.github.com/repos/{REPO}/contents/{path}?ref=main", token)
    sha = meta['sha']
    content = open(local_path, 'rb').read()
    b64 = base64.b64encode(content).decode()
    payload = json.dumps({"message": msg, "content": b64, "sha": sha, "branch": "main"}).encode()
    result = _req("PUT", f"https://api.github.com/repos/{REPO}/contents/{path}", token, payload)
    print(path, "-> new sha:", result['content']['sha'])
    return result


def verify(path, token):
    """배포 후 실제 반영 확인용. web_fetch 대신 이걸 쓸 것."""
    d = _req("GET", f"https://api.github.com/repos/{REPO}/contents/{path}?ref=main", token)
    return base64.b64decode(d['content']).decode('utf-8')


def check_pages_build(token):
    b = _req("GET", f"https://api.github.com/repos/{REPO}/pages/builds/latest", token)
    return b.get('status'), b.get('error')


def nudge_rebuild(html_path, token):
    """Pages 빌드 큐가 stuck/errored일 때: HTML 파일에 빈 주석 한 줄 추가하는 no-op 커밋으로 재트리거."""
    import datetime
    meta = _req("GET", f"https://api.github.com/repos/{REPO}/contents/{html_path}?ref=main", token)
    sha = meta['sha']
    content = base64.b64decode(meta['content']).decode('utf-8')
    content += f"\n<!-- rebuild-nudge {datetime.datetime.now(datetime.UTC).isoformat()} -->\n"
    b64 = base64.b64encode(content.encode('utf-8')).decode()
    payload = json.dumps({"message": "no-op: trigger pages rebuild", "content": b64, "sha": sha, "branch": "main"}).encode()
    result = _req("PUT", f"https://api.github.com/repos/{REPO}/contents/{html_path}", token, payload)
    print("nudge commit sha:", result['content']['sha'])


if __name__ == '__main__':
    print("이 파일은 import해서 쓰는 헬퍼입니다. 예: from deploy import put_file, verify, check_pages_build")
