"""Build index.html from src/ (one self-contained page). Run: python3 build.py"""
from pathlib import Path

src = Path(__file__).parent / "src"
html = (src / "template.html").read_text()
css = (src / "style.css").read_text()
js = "\n".join((src / f).read_text() for f in ["engine.js", "labs_bc.js", "keyassess.js", "minigrid.js", "ui.js"])
html = html.replace("/*CSS*/", css).replace("/*JS*/", js)
(Path(__file__).parent / "index.html").write_text(html)
print("index.html written,", len(html), "bytes")
