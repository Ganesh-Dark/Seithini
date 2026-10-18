# Seithini (செய்தினி)

Seithini collects daily news from major Tamil and English newspapers across Tamil Nadu, organizes them in one place, and helps you stay updated without opening multiple websites.

---

## Why Seithini?

Checking news across different Tamil Nadu newspapers usually means visiting multiple websites or apps one by one. Seithini pulls the latest stories from top dailies into a single clean list in just a few seconds.

### Newspapers Covered:
- **Tamil:** Daily Thanthi, Dinamani, Maalai Malar, Dinamalar
- **English:** The Hindu, The New Indian Express, DT Next

---

## How It Works

Instead of scraping web pages that break often, Seithini connects directly to the official feeds and APIs used by these newspapers:

| Newspaper | Language | Source Type |
| :--- | :--- | :--- |
| **Daily Thanthi** | Tamil | Official JSON API |
| **Dinamani** | Tamil | Official JSON API |
| **Maalai Malar** | Tamil | Official JSON API |
| **Dinamalar** | Tamil | News Feed |
| **The Hindu** | English | News Feed |
| **The New Indian Express** | English | Official JSON API |
| **DT Next** | English | Official JSON API |

Each collected story includes the headline, a brief summary, the publication date, and a direct link to read the original article on the publisher's site.

---

## Quick Start

Make sure you have `uv` installed, then run:

```bash
uv run generate_news_excel.py
```

This collects around 350+ fresh stories across all 7 newspapers in a few seconds.

---

## What We Are Building Next

1. **Daily Archive:** Save each day's news into a simple database so past news is kept safe and duplicates are ignored.
2. **Topic Grouping:** Automatically sort stories into everyday topics like Politics, Cinema, Weather, and Sports.
3. **Ask Seithini (AI Chat):** A helpful chatbot where you can ask everyday questions like *"What happened in Chennai today?"* or *"இன்றைய மழை நிலவரம் என்ன?"* in plain Tamil or English and get direct answers with links to the original papers.
4. **Website:** Host this on a clean website with a custom domain so anyone can read it easily from their phone or laptop.
