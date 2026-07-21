# 代码块效果演示

这是一段普通的正文文字，用来对比代码块与正文的视觉差异。
可以看到代码块有浅灰色背景、深灰色边框和等宽字体。

## 一个 Python 爬虫示例

下面这段代码展示了表格包裹后的完整效果：

```python
import requests
from bs4 import BeautifulSoup

def fetch_titles(url: str) -> list[str]:
    """Fetch all h2 titles from a given URL."""
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')
    return [h2.get_text(strip=True) for h2 in soup.find_all('h2')]

if __name__ == '__main__':
    titles = fetch_titles('https://example.com')
    for t in titles:
        print(f' - {t}')
```

可以看到代码块与正文之间有 6pt 的间距，边框是精致的 0.5pt 深灰线，
文字与边框之间有 4pt 的内边距。每行代码使用 Consolas 9pt 等宽字体，
行距固定为 12pt，整体紧凑而清晰。

## 一个短代码块

```bash
docker compose up -d --build
docker compose logs -f
```

短代码块同样有完整的背景、边框和内边距，视觉效果一致。
