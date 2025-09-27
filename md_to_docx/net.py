from io import BytesIO
import requests
from .errors import ImageDownloadError

class HttpClient:
    def __init__(self, *, timeout=(5,15), max_bytes=10*1024*1024, user_agent="md-to-docx/1.0"):
        self.timeout = timeout
        self.max_bytes = max_bytes
        self.headers = {"User-Agent": user_agent}

    def fetch_image(self, url: str) -> BytesIO:
        try:
            with requests.get(url, stream=True, timeout=self.timeout, headers=self.headers) as r:
                r.raise_for_status()
                ctype = r.headers.get("Content-Type", "")
                if not ctype.startswith("image/"):
                    raise ImageDownloadError(f"{url}: Content-Type '{ctype}' ist kein Bild.")
                buf = BytesIO(); size = 0
                for chunk in r.iter_content(8192):
                    if not chunk: continue
                    size += len(chunk)
                    if size > self.max_bytes:
                        raise ImageDownloadError(f"{url}: Bild größer als {self.max_bytes} Bytes.")
                    buf.write(chunk)
                buf.seek(0); return buf
        except requests.RequestException as e:
            raise ImageDownloadError(f"{url}: {e}") from e
