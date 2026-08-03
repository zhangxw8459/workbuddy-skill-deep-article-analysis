# -*- coding: utf-8 -*-
"""截图脚本：将报告 HTML 渲染为长图与封面图，供 README 展示"""
import asyncio
import pathlib
from playwright.async_api import async_playwright

DEMO_DIR = pathlib.Path(__file__).resolve().parent
HTML_FILE = DEMO_DIR / "2026-08-03_牛市最重要的是下血本重仓.html"


async def main():
    html_url = HTML_FILE.as_uri()
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1200, "height": 800})
        await page.goto(html_url, wait_until="networkidle")
        await page.screenshot(path=str(DEMO_DIR / "demo-report-full.png"), full_page=True)
        await page.screenshot(path=str(DEMO_DIR / "demo-report-cover.png"))
        # 只截正文中间一段（内容板块），便于 README 展示 10 板块结构
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.25)")
        await page.wait_for_timeout(300)
        await page.screenshot(path=str(DEMO_DIR / "demo-report-content.png"))
        await browser.close()
        print("截图完成")


asyncio.run(main())
