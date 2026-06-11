"""
국가법령정보센터 Open API 클라이언트
- 재시도 3회 (대기 시간 점진적 증가)
- 요청 간 0.5초 딜레이 (API 제한 대응)
- UTF-8 → EUC-KR 인코딩 fallback
"""
import time
import requests

BASE_SEARCH_URL = "http://www.law.go.kr/DRF/lawSearch.do"
BASE_SERVICE_URL = "http://www.law.go.kr/DRF/lawService.do"

MAX_RETRIES = 3
REQUEST_DELAY = 0.5   # 요청 간 최소 간격 (초)
REQUEST_TIMEOUT = 30


class APIError(Exception):
    pass


class LawAPIClient:
    def __init__(self, api_key: str) -> None:
        self.api_key = api_key
        self.session = requests.Session()
        self._last_request_time: float = 0

    def _wait(self) -> None:
        elapsed = time.time() - self._last_request_time
        if elapsed < REQUEST_DELAY:
            time.sleep(REQUEST_DELAY - elapsed)

    def _decode(self, response: requests.Response) -> str:
        """UTF-8 우선, 실패 시 EUC-KR로 디코딩."""
        try:
            return response.content.decode("utf-8")
        except UnicodeDecodeError:
            return response.content.decode("euc-kr")

    def _get(self, url: str, params: dict) -> str:
        params = {**params, "OC": self.api_key, "type": "XML"}
        self._wait()

        for attempt in range(MAX_RETRIES):
            try:
                resp = self.session.get(url, params=params, timeout=REQUEST_TIMEOUT)
                resp.raise_for_status()
                self._last_request_time = time.time()
                return self._decode(resp)
            except requests.RequestException as e:
                if attempt < MAX_RETRIES - 1:
                    wait = attempt + 1  # 1초, 2초, 3초
                    print(f"  ⚠  요청 실패 (시도 {attempt + 1}/{MAX_RETRIES}): {e}. {wait}초 후 재시도...")
                    time.sleep(wait)
                else:
                    raise APIError(f"API 요청 최종 실패: {e}") from e

    def search_law(self, query: str) -> str:
        """법령명으로 법령 검색 (target=law)."""
        return self._get(BASE_SEARCH_URL, {"target": "law", "query": query, "display": 20, "page": 1})

    def search_3stage_laws(self, query: str) -> str:
        """3단비교 법령 목록 검색 (target=thdCmp).
        법률·시행령·시행규칙 MST를 한번에 반환."""
        return self._get(BASE_SEARCH_URL, {"target": "thdCmp", "query": query, "display": 20, "page": 1})

    def get_3stage_comparison(self, mst: str) -> str:
        """위임조문 3단비교표 조회 (target=thdCmp, knd=2). MST = 법령일련번호."""
        return self._get(BASE_SERVICE_URL, {"target": "thdCmp", "MST": mst, "knd": 2})

    def get_law_articles(self, mst: str) -> str:
        """법령 전체 조문 조회 - 갭 분석용 (target=law). 파라미터명 MST = 법령일련번호."""
        return self._get(BASE_SERVICE_URL, {"target": "law", "MST": mst})
