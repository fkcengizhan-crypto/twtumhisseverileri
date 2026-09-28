#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TradingView Hisse Tarama → Excel
================================

TradingView'in "Sütun ekle" listesindeki tüm metrikleri otomatik keşfeder,
seçilen piyasadaki hisselerin verilerini çeker ve tek bir Excel dosyasına yazar.

Nasıl çalışır
--------------
1. Screener sayfasındaki JavaScript paketlerini tarar; "Sütun ekle" listesinin
   tanımları bu paketlerin içindedir (bir API değil, frontend kodu).
2. Sütun adaylarını parametrelere göre genişletir (mali dönem, gösterge uzunluğu,
   zaman dilimi, plot, büyüme tipi).
3. Adayları TradingView'in scanner uç noktasına sorar ve verisi olmayanları eler.
4. Kalan verileri tek bir Excel dosyasına yazar.

Gereksinimler
-------------
    Python 3.9+  ve  openpyxl

    pip install openpyxl

Kullanım
--------
    python tradingview_screener.py                      # BIST, tüm sütunlar
    python tradingview_screener.py --katalog-yenile     # sütunları yeniden keşfet
    python tradingview_screener.py --sadece-hisse       # fon/ETF'leri çıkar
    python tradingview_screener.py --sutun "close,Perf.Y,RSI"
    python tradingview_screener.py --market usa
    python tradingview_screener.py --katalog-yaz sutun_katalogu.csv

Excel çıktısı
-------------
    1. sayfa  BIST verisi. Üst satır Türkçe etiket, alt satır TradingView anahtarı.
              İlk iki sütun ve ilk iki satır dondurulmuş, AutoFilter açık.
    2. sayfa  Sütun Kataloğu: anahtar, etiket, kategori, dolu satır sayısı.

Etiket değiştirmek için LABEL_MAP ve SECTOR_MAP sözlüklerini düzenleyip
betiği tekrar çalıştırmak yeterlidir.

Çıktı dizini
------------
    Her çalıştırma "Archive" klasörüne üç dosya yazar (betiği çalıştırdığınız
    klasörün altında):
        PİYASA_YYYYMMDD.xlsx            tarihli Excel
        Bist_Hisse_Tum_Veriler.xlsx     sabit isimli kopya (her seferinde yenilenir)
        Bist_Hisse_Tum_Veriler.json     aynı verinin JSON hâli, satır tabanlı

    Sütun önbelleği ise kullanıcı klasöründeki ".tradingview_screener" içinde
    tutulur; TRV_HOME ile değiştirilebilir.

Not
---
Bu araç TradingView'in resmî olmayan, herkese açık tarayıcı uç noktasını
kullanır. Login gerektirmez, istekler arası kısa beklemeler vardır. Görülen
tüm veriler TradingView'a aittir; kullanım şartlarına uymak sizindir.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import urllib.request

VERSION = "1.0.0"

# Sütun önbelleği ve çıktı dizini. TRV_HOME ortam değişkeniyle önbellek yolu
# değiştirilebilir; Excel dosyaları her zaman çalıştırılan klasörün altına gider.
HOME_DIR = Path(os.environ.get("TRV_HOME") or Path.home() / ".tradingview_screener")
CATALOG_FILE = HOME_DIR / "sutun_katalogu.json"
OUT_DIR = Path.cwd() / "Archive"

SCANNER_URL = "https://scanner.tradingview.com/{market}/scan"
SCREENER_PAGE = "https://tr.tradingview.com/screener/"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Origin": "https://tr.tradingview.com",
    "Referer": "https://tr.tradingview.com/",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

TIMEOUT = 60
RETRIES = 3
BATCH = 150
LANG = "tr"

MARKET_NAMES = {
    "turkey": "BIST", "bist": "BIST", "turkeycrypto": "BIST Kripto",
    "usa": "ABD", "germany": "Almanya", "japan": "Japonya", "uk": "İngiltere",
    "france": "Fransa", "india": "Hindistan", "korea": "Güney Kore",
    "hongkong": "Hong Kong", "brazil": "Brezilya", "canada": "Kanada",
    "italy": "İtalya", "spain": "İspanya", "sweden": "İsveç", "norway": "Norveç",
    "switzerland": "İsviçre", "australia": "Avustralya", "netherlands": "Hollanda",
    "poland": "Polonya", "russia": "Rusya", "saudiarabia": "S. Arabistan",
    "taiwan": "Tayvan", "singapore": "Singapur", "southafrica": "G. Afrika",
    "global": "Global",
}

# -------------------------------------------------------------------------
# Turkce etiket haritasi (anahtar -> etiket)
# Kendi etiketinizi eklemek için: anahtarı aynen yazın, değeri değiştirin.
# Önerilen etiketler Excel'de 58 karakterden uzun olmamalıdır.
# -------------------------------------------------------------------------
LABEL_MAP = {
    'name': 'Hisse Adı',
    'description': 'Şirket Tam Adı',
    'sector': 'Sektör',
    'industry': 'Alt Sektör',
    'exchange': 'Borsa',
    'type': 'Tür',
    'currency': 'Para Birimi',
    'close': 'Son Fiyat',
    'open': 'Açılış',
    'high': 'En Yüksek',
    'low': 'En Düşük',
    'change': 'Günlük Değişim (%)',
    'price_52_week_high': '52 Haftalık En Yüksek',
    'price_52_week_low': '52 Haftalık En Düşük',
    'high_price_52w': '52 Haftalık En Yüksek',
    'low_price_52w': '52 Haftalık En Düşük',
    'country': 'Ülke',
    'isin': 'ISIN (Hisse Kimlik No)',
    'volume': 'Hacim',
    'relative_volume_10d_calc': 'Göreli Hacim (10 Günlük)',
    'relative_volume_30d_calc': 'Göreli Hacim (30 Günlük)',
    'average_volume': 'Ortalama Hacim',
    'average_volume_10': 'Ortalama Hacim (10 Günlük)',
    'average_volume_30': 'Ortalama Hacim (30 Günlük)',
    'average_volume_60': 'Ortalama Hacim (60 Günlük)',
    'average_volume_10_calc': 'Ortalama Hacim (10 Günlük)',
    'average_volume_20_calc': 'Ortalama Hacim (20 Günlük)',
    'average_volume_30_calc': 'Ortalama Hacim (30 Günlük)',
    'average_volume_60_calc': 'Ortalama Hacim (60 Günlük)',
    'average_volume_90_calc': 'Ortalama Hacim (90 Günlük)',
    'average_volume_120_calc': 'Ortalama Hacim (120 Günlük)',
    'average_volume_180_calc': 'Ortalama Hacim (180 Günlük)',
    'average_volume_240_calc': 'Ortalama Hacim (240 Günlük)',
    'average_volume_10d_calc': 'Ortalama Hacim (10 Günlük)',
    'average_volume_30d_calc': 'Ortalama Hacim (30 Günlük)',
    'market_cap_basic': 'Piyasa Değeri',
    'market_cap_calc': 'Piyasa Değeri (Hesaplanan)',
    'market_cap_diluted_calc': 'Piyasa Değeri (Seyreltilmiş)',
    'enterprise_value_ebitda_ttm': 'Firma Değeri / FAVÖK (TTM)',
    'enterprise_value': 'Firma Değeri',
    'enterprise_value_fq': 'Firma Değeri (Çeyreklik)',
    'price_earnings_ttm': 'F/K (TTM)',
    'price_earnings_forward': 'İleriye Dönük F/K',
    'price_book_fq': 'F/D (Çeyreklik)',
    'price_book_forward': 'İleriye Dönük F/D',
    'price_sales_current': 'F/Satışlar',
    'price_sales_ratio_current': 'F/Satışlar',
    'price_to_cash_ratio': 'F/Nakit',
    'price_to_free_cash_flow': 'F/Serbest Nakit Akışı',
    'peg_ratio': 'PEG Oranı',
    'beta_1_year': 'Beta (1 Yıllık)',
    'beta_3_year': 'Beta (3 Yıllık)',
    'beta_5_year': 'Beta (5 Yıllık)',
    'dividend_yield_recent': 'Temettü Verimi (%)',
    'payout_ratio': 'Temettü Dağıtım Oranı (%)',
    'dividends_per_share_fq': 'Hisse Başına Temettü (Çeyreklik)',
    'dividends_per_share': 'Hisse Başına Temettü',
    'dividend_per_share_recent': 'Hisse Başına Temettü (Son)',
    'earnings_per_share_basic_fy': 'Hisse Başına Kazanç (Yıllık)',
    'earnings_per_share_basic_ttm': 'Hisse Başına Kazanç (TTM)',
    'earnings_per_share_diluted_ttm': 'Hisse Başına Kazanç (Seyreltilmiş',
    'earnings_per_share_basic': 'Hisse Başına Kazanç',
    'book_value_per_share_fq': 'Hisse Başına Defter Değeri (Çeyreklik)',
    'book_value_per_share': 'Hisse Başına Defter Değeri',
    'total_revenue': 'Toplam Gelir',
    'total_revenue_fy': 'Toplam Gelir (Yıllık)',
    'total_revenue_fq': 'Toplam Gelir (Çeyreklik)',
    'total_revenue_ttm': 'Toplam Gelir (TTM)',
    'total_revenue_yoy_growth_fy': 'Toplam Gelir Büyüme (Yıllık',
    'total_revenue_yoy_growth_ttm': 'Toplam Gelir Büyüme (TTM',
    'total_revenue_cagr_5y': 'Toplam Gelir 5 Yıllık Bileşik Büyüme (%)',
    'gross_profit': 'Brüt Kâr',
    'gross_profit_fy': 'Brüt Kâr (Yıllık)',
    'gross_profit_fq': 'Brüt Kâr (Çeyreklik)',
    'gross_profit_ttm': 'Brüt Kâr (TTM)',
    'operating_income': 'Faaliyet Kârı',
    'operating_income_fy': 'Faaliyet Kârı (Yıllık)',
    'operating_income_fq': 'Faaliyet Kârı (Çeyreklik)',
    'net_income': 'Net Kâr',
    'net_income_fy': 'Net Kâr (Yıllık)',
    'net_income_fq': 'Net Kâr (Çeyreklik)',
    'net_income_ttm': 'Net Kâr (TTM)',
    'net_income_common_stockholders': 'Ortaklara Net Kâr',
    'net_income_yoy_growth_fy': 'Net Kâr Büyüme (Yıllık',
    'net_income_yoy_growth_ttm': 'Net Kâr Büyüme (TTM',
    'net_income_cagr_5y': 'Net Kâr 5 Yıllık Bileşik Büyüme (%)',
    'ebitda': 'FAVÖK',
    'ebitda_fq': 'FAVÖK (Çeyreklik)',
    'ebitda_ttm': 'FAVÖK (TTM)',
    'ebit': 'FAVIT',
    'ebit_fq': 'FAVIT (Çeyreklik)',
    'rnd': 'Ar-Ge Harcaması',
    'rnd_ratio': 'Ar-Ge / Gelir (%)',
    'total_assets': 'Toplam Varlıklar',
    'total_assets_fq': 'Toplam Varlıklar (Çeyreklik)',
    'total_liabilities': 'Toplam Yükümlülükler',
    'total_equity': 'Toplam Özkaynak',
    'shareholders_equity': 'Özkaynak (Hissedarlara Ait)',
    'cash_and_equivalents': 'Nakit ve Nakit Benzeri',
    'cash_and_short_term_investments': 'Kısa Vadeli Yatırımlar',
    'total_debt': 'Toplam Borç',
    'net_debt': 'Net Borç',
    'working_capital': 'İşletme Sermayesi',
    'invested_capital': 'Yatırılmış Sermaye',
    'free_cash_flow': 'Serbest Nakit Akışı',
    'free_cash_flow_fy': 'Serbest Nakit Akışı (Yıllık)',
    'operating_cash_flow': 'Faaliyet Nakit Akışı',
    'operating_cash_flow_fy': 'Faaliyet Nakit Akışı (Yıllık)',
    'capex': 'Seraye Yatırım',
    'capex_fy': 'Seraye Yatırım (Yıllık)',
    'depreciation_amortization': 'Amortisman',
    'gross_margin': 'Brüt Kar (%)',
    'operating_margin': 'Faaliyet Kar Marjı (%)',
    'net_margin': 'Net Kar Marjı (%)',
    'pre_tax_margin': 'V Öncesi Kar Marjı (%)',
    'return_on_equity': 'Özkaynak Getirisi (ROE %)',
    'return_on_equity_fq': 'Özkaynak Getirisi (ROE %',
    'return_on_assets': 'Varlık Getirisi (ROA %)',
    'return_on_invested_capital': 'Yatırılmış Sermaye Getirisi (ROIC %)',
    'current_ratio': 'Current Ratio',
    'quick_ratio': 'Quick Ratio',
    'debt_to_equity': 'Borç / Özkaynak',
    'debt_to_equity_fq': 'Borç / Özkaynak (Çeyreklik)',
    'float_shares_outstanding': 'Halka Açık Hisse Adedi',
    'shares_outstanding': 'Dolaşımdaki Hisse Adedi',
    'number_of_employees': 'Çalışan Sayısı',
    'revenue_per_employee': 'Çalışan Başına Gelir',
    'net_income_per_employee': 'Çalışan Başına Kâr',
    'recommend_all': 'Teknik Tavsiye (Genel)',
    'recommend_ma': 'Teknik Tavsiye (Hareketli Ortalamalar)',
    'recommend_other': 'Teknik Tavsiye (Diğer)',
    'Recommend.All': 'Teknik Tavsiye (Genel)',
    'Recommend.MA': 'Teknik Tavsiye (Hareketli Ortalamalar)',
    'Recommend.Other': 'Teknik Tavsiye (Diğer)',
    'recommendation_mark': 'Teknik Tavsiye Puanı',
    'Volatility.D': 'Günlük Volatilite (%)',
    'Volatility.W': 'Haftalık Volatilite (%)',
    'Volatility.M': 'Aylık Volatilite (%)',
    'dividends_yield': 'Temettü Verimi (%)',
    'analyst_rating': 'Analist Tavsiyesi',
    'number_of_analysts': 'Analist Sayısı',
    'target_price_average': 'Hedef Fiyat (Ortalama)',
    'target_price_median': 'Hedef Fiyat (Medyan)',
    'target_price_high': 'Hedef Fiyat (En Yüksek)',
    'target_price_low': 'Hedef Fiyat (En Düşük)',
    'earnings_release_date': 'Bir Sonraki Finansal Rapor Tarihi',
    'earnings_release_next_date': 'Finansal Rapor Tarihi',
    'earnings_per_share_forecast': 'EPS Tahmini',
    'earnings_per_share_forecast_next_fq': 'EPS Tahmini (Gelecek Çeyrek)',
    'earnings_per_share_forecast_next_fy': 'EPS Tahmini (Gelecek Yıl)',
    'earnings_per_share_forecast_fq': 'EPS Tahmini (Bu Çeyrek)',
    'earnings_per_share_forecast_fy': 'EPS Tahmini (Bu Yıl)',
    'Perf.W': 'Haftalık Performans',
    'Perf.1M': '1 Aylık Performans',
    'Perf.3M': '3 Aylık Performans',
    'Perf.6M': '6 Aylık Performans',
    'Perf.Y': '1 Yıllık Performans',
    'Perf.2Y': '2 Yıllık Performans',
    'Perf.5Y': '5 Yıllık Performans',
    'Perf.10Y': '10 Yıllık Performans',
    'Perf.All': 'Tüm Zamanlar Performans',
    'change|1': 'Değişim (1 Gün)',
    'change|5': 'Değişim (5 Gün)',
    'change|15': 'Değişim (15 Gün)',
    'change|30': 'Değişim (1 Ay)',
    'change|60': 'Değişim (2 Ay)',
    'change|1W': 'Değişim (1 Hafta)',
    'change|1M': 'Değişim (1 Ay)',
    'change|3M': 'Değişim (3 Ay)',
    'change|6M': 'Değişim (6 Ay)',
    'change|1Y': 'Değişim (1 Yıl)',
    'change|5Y': 'Değişim (5 Yıl)',
    'change|YTD': 'Yıl Başından beri Değişim',
    'change|All': 'Tüm Zamanlar Değişim',
    'High.1M': '1 Aylık En Yüksek',
    'Low.1M': '1 Aylık En Düşük',
    'High.3M': '3 Aylık En Yüksek',
    'Low.3M': '3 Aylık En Düşük',
    'High.6M': '6 Aylık En Yüksek',
    'Low.6M': '6 Aylık En Düşük',
    'High.1Y': '1 Yıllık En Yüksek',
    'Low.1Y': '1 Yıllık En Düşük',
    'High.All': 'Tüm Zamanlar En Yüksek',
    'Low.All': 'Tüm Zamanlar En Düşük',
    'RSI': 'RSI (14)',
    'RSI|1': 'RSI (Günlük Kapanış)',
    'RSI|5': 'RSI (5 Dakikalık)',
    'RSI|15': 'RSI (15 Dakikalık)',
    'RSI|30': 'RSI (30 Dakikalık)',
    'RSI|60': 'RSI (1 Saatlik)',
    'RSI|240': 'RSI (4 Saatlik)',
    'RSI|1D': 'RSI (Günlük)',
    'RSI|1W': 'RSI (Haftalık)',
    'MACD.macd': 'MACD',
    'MACD.signal': 'MACD Sinyal',
    'Stoch.K': 'Stochastic %K',
    'Stoch.D': 'Stochastic %D',
    'Stoch.RSI.K': 'Stochastic RSI %K',
    'Stoch.RSI.D': 'Stochastic RSI %D',
    'CCI': 'CCI (20)',
    'ADX': 'ADX (14)',
    'ADX+DI': 'ADX +DI (14)',
    'ADX-DI': 'ADX -DI (14)',
    'ATR': 'ATR (14)',
    'BB.upper': 'Bollinger Üst Bandı',
    'BB.lower': 'Bollinger Alt Bandı',
    'BBPower': 'Bollinger Bant Gücü',
    'BB.middle': 'Bollinger Orta Band',
    'Mom': 'Momentum (10)',
    'AO': 'Aroon Osilatör',
    'UO': 'Ultimate Osilatör',
    'W.R': 'Williams %R',
    'P.SAR': 'Parabolik SAR',
    'Aroon': 'Aroon (14)',
    'Aroon.up': 'Aroon Yukarı',
    'Aroon.dn': 'Aroon Aşağı',
    'Ichimoku': 'Ichimoku Bulutu',
    'Ichimoku.BLine': 'Ichimoku Temel Çizgi',
    'VWAP': 'VWAP',
    'VWMA': 'VWMA (20)',
    'SMA20': 'SMA (20)',
    'SMA50': 'SMA (50)',
    'SMA200': 'SMA (200)',
    'EMA20': 'EMA (20)',
    'EMA50': 'EMA (50)',
    'EMA200': 'EMA (200)',
    'MoneyFlow.index': 'Money Flow Index',
    'ChaikinMoneyFlow': 'Chaikin Para Akışı',
    'TSI': 'TSI',
    'Donchian.middle': 'Donchian Orta',
    'KltChnl.middle': 'Keltner Kanal Orta',
    'Pivot.M.Classic.S1': 'Klasik Pivot S1',
    'Pivot.M.Classic.R1': 'Klasik Pivot R1',
    'Pivot.M.Fibonacci.S1': 'Fibonacci Pivot S1',
    'Pivot.M.Fibonacci.R1': 'Fibonacci Pivot R1',
    'ATR|1D': 'ATR (Günlük)',
}


# Sektör / alt sektör / ülke çevirisi. Yeni değer ekleyebilirsin.
SECTOR_MAP = {
    'Turkey': 'Türkiye',
    'United States': 'ABD',
    'Germany': 'Almanya',
    'Finance': 'Finans',
    'Process Industries': 'İşlem Sanayii',
    'Producer Manufacturing': 'Üretim Sanayi',
    'Consumer Non-Durables': 'Tüketim Ürünleri (Dayanıksız)',
    'Non-Energy Minerals': 'Enerji Dışı Madencilik',
    'Utilities': 'Kamu Hizmetleri',
    'Consumer Services': 'Tüketim Hizmetleri',
    'Consumer Durables': 'Tüketim Ürünleri (Dayanıklı)',
    'Miscellaneous': 'Diğer',
    'Technology Services': 'Teknoloji Hizmetleri',
    'Distribution Services': 'Dağıtım Hizmetleri',
    'Retail Trade': 'Perakende',
    'Industrial Services': 'Endüstriyel Hizmetler',
    'Electronic Technology': 'Elektronik Teknoloji',
    'Commercial Services': 'Ticari Hizmetler',
    'Transportation': 'Ulaşım',
    'Health Technology': 'Sağlık Teknolojisi',
    'Health Services': 'Sağlık Hizmetleri',
    'Energy Minerals': 'Enerji Madenciliği',
    'Communications': 'İletişim',
    'Real Estate Investment Trusts': 'Gayrimenkul Yatırım Ortaklıkları',
    'Food: Specialty/Candy': 'Gıda: Özel Ürünler / Şekerleme',
    'Real Estate Development': 'Gayrimenkul Geliştirme',
    'Construction Materials': 'İnşaat Malzemeleri',
    'Textiles': 'Tekstil',
    'Steel': 'Çelik',
    'Alternative Power Generation': 'Yenilenebilir Enerji Üretimi',
    'Investment Managers': 'Yatırım Yöneticileri',
    'Investment Trusts/Mutual Funds': 'Yatırım Fonları / Borsa Yatırım Fonları',
    'Finance/Rental/Leasing': 'Finans / Kiralama / Leasing',
    'Electric Utilities': 'Elektrik Dağıtım Şirketleri',
    'Agricultural Commodities/Milling': 'Tarımsal Ürünler / Değirmencilik',
    'Engineering & Construction': 'Mühendislik ve İnşaat',
    'Containers/Packaging': 'Ambalaj',
    'Electrical Products': 'Elektrikli Ürünler',
    'Investment Banks/Brokers': 'Yatırım Bankaları / Aracı Kurumlar',
    'Packaged Software': 'Paket Yazılım',
    'Industrial Specialties': 'Endüstriyel Ürünler',
    'Apparel/Footwear': 'Giyim / Ayakkabı',
    'Industrial Machinery': 'Endüstriyel Makinalar',
    'Hotels/Resorts/Cruise lines': 'Otel / Turizm / Gemi Seyahatleri',
    'Financial Conglomerates': 'Finans Holdingleri',
    'Miscellaneous Manufacturing': 'Diğer İmalat',
    'Major Banks': 'Büyük Bankalar',
    'Wholesale Distributors': 'Toptancı Dağıtıcılar',
    'Auto Parts: OEM': 'Otomotiv Yan Sanayi (OEM)',
    'Electronics Distributors': 'Elektronik Dağıtımcılar',
    'Building Products': 'Yapı Ürünleri',
    'Food Retail': 'Gıda Perakende',
    'Home Furnishings': 'Ev Tekstili / Mobilya',
    'Food: Meat/Fish/Dairy': 'Gıda: Et / Balık / Süt Ürünleri',
    'Food: Major Diversified': 'Gıda: Çeşitlendirilmiş',
    'Chemicals: Specialty': 'Kimyasallar: Özel Ürünler',
    'Miscellaneous Commercial Services': 'Diğer Ticari Hizmetler',
    'Trucks/Construction/Farm Machinery': 'İş / İnşaat / Tarım Makinaları',
    'Metal Fabrication': 'Metal İşleme',
    'Information Technology Services': 'Bilgi Teknolojisi Hizmetleri',
    'Beverages: Non-Alcoholic': 'İçecek: Alkolsüz',
    'Other Consumer Services': 'Diğer Tüketici Hizmetleri',
    'Restaurants': 'Restoranlar',
    'Specialty Stores': 'Özel Mağazalar',
    'Pharmaceuticals: Major': 'İlaçlar: Büyük Ölçekli',
    'Movies/Entertainment': 'Sinema / Eğlence',
    'Gas Distributors': 'Gaz Dağıtım',
    'Regional Banks': 'Bölgesel Bankalar',
    'Other Metals/Minerals': 'Diğer Metaller / Madenler',
    'Commercial Printing/Forms': 'Ticari Baskı / Formlar',
    'Household/Personal Care': 'Ev ve Kişisel Bakım',
    'Multi-Line Insurance': 'Çok Branşlı Sigorta',
    'Motor Vehicles': 'Motorlu Araçlar',
    'Electronics/Appliances': 'Elektronik / Beyaz Eşya',
    'Aerospace & Defense': 'Havacılık ve Savunma',
    'Air Freight/Couriers': 'Hava Kargo / Kurye',
    'Hospital/Nursing Management': 'Hastane / Bakımevi Yönetimi',
    'Chemicals: Agricultural': 'Kimyasallar: Tarımsal',
    'Other Transportation': 'Diğer Ulaşım',
    'Medical Specialties': 'Tıbbi Uzmanlık',
    'Marine Shipping': 'Denizcilik',
    'Electronic Equipment/Instruments': 'Elektronik Cihazlar ve Aletler',
    'Pulp & Paper': 'Kâğıt ve Hamur',
    'Chemicals: Major Diversified': 'Kimyasallar: Çeşitlendirilmiş',
    'Life/Health Insurance': 'Hayat / Sağlık Sigortası',
    'Aluminum': 'Alüminyum',
    'Industrial Conglomerates': 'Endüstriyel Holdingler',
    'Biotechnology': 'Biyoteknoloji',
    'Electronic Production Equipment': 'Elektronik Üretim Ekipmanları',
    'Oil Refining/Marketing': 'Petrol Rafinerisi / Pazarlama',
    'Precious Metals': 'Kıymetli Metaller',
    'Other Consumer Specialties': 'Diğer Tüketici Ürünleri',
    'Medical/Nursing Services': 'Tıbbi / Bakım Hizmetleri',
    'Forest Products': 'Orman Ürünleri',
    'Publishing: Newspapers': 'Yayıncılık: Gazeteler',
    'Automotive Aftermarket': 'Otomotiv Yedek Parça',
    'Airlines': 'Havayolları',
    'Computer Communications': 'Bilgisayar İletişim',
    'Semiconductors': 'Yarı İletkenler',
    'Advertising/Marketing Services': 'Reklam / Pazarlama Hizmetleri',
    'Major Telecommunications': 'Büyük Telekom',
    'Office Equipment/Supplies': 'Büro Ekipmanları',
    'Telecommunications Equipment': 'Telekomünikasyon Ekipmanları',
    'Recreational Products': 'Boş Zaman Ürünleri',
    'Food Distributors': 'Gıda Dağıtımcıları',
    'Consumer Sundries': 'Tüketim Çeşitli Ürünler',
    'Department Stores': 'Mağazalar',
    'Internet Retail': 'İnternet Perakende',
    'Data Processing Services': 'Veri İşleme Hizmetleri',
    'Trucking': 'Karayolu Taşımacılığı',
    'Property/Casualty Insurance': 'Gayrimenkul / Kaza Sigortası',
    'Homebuilding': 'Konut İnşaatı',
    'Electronics/Appliance Stores': 'Elektronik / Beyaz Eşya Mağazaları',
    'Apparel/Footwear Retail': 'Giyim / Ayakkabı Perakende',
    'Wireless Telecommunications': 'Kablosuz Telekomünikasyon',
    'Media Conglomerates': 'Medya Holdingleri',
    'Railroads': 'Demiryolları',
    'Pharmaceuticals: Other': 'İlaçlar: Diğer',
    'Publishing: Books/Magazines': 'Yayıncılık: Kitap / Dergi',
    'Integrated Oil': 'Entegre Petrol',
    'Beverages: Alcoholic': 'İçecek: Alkollü',
    'Medical Distributors': 'Tıbbi Dağıtımcılar',
    'Computer Processing Hardware': 'Bilgisayar İşlem Donanımı',
}


# --------------------------------------------------------------------------
# Dogrulanmis isimlendirme kurallari
#
# Bu kalipler TradingView'in kendi frontend paketinden cikarildi ve
# scanner API'sine sorularak tek tek teyit edildi. Yanlis olanlar
# (orn. "Perf.1Y", "ADX|14", "total_revenue_annual") bilerek kullanilmaz.
# --------------------------------------------------------------------------

FISCAL_SUFFIXES = ["_fy", "_fq", "_ttm"]
GROWTH_SUFFIXES = [
    "_yoy_growth_ttm", "_yoy_growth_fy", "_yoy_growth_fq",
    "_qoq_growth_fq", "_cagr_5y",
]
LENGTHS = [5, 9, 10, 12, 14, 20, 21, 26, 30, 50, 100, 200, 250]
CALC_LENGTHS = [10, 20, 30, 60, 90, 120, 180, 240, 250]
# gösterge zaman dilimleri: |1D |3M |6M |1Y geçerli DEĞİL
TF_SUFFIXES = ["|1", "|5", "|15", "|30", "|60", "|120", "|240", "|1W", "|1M"]
# "change" yalnızca sayılı gün kombinasyonlarını ve 1W/1M kabul ediyor
CHANGE_SUFFIXES = ["|1", "|5", "|15", "|30", "|60", "|120", "|240", "|1W", "|1M"]
PERF_PERIODS = ["W", "1M", "3M", "6M", "Y", "3Y", "5Y", "10Y", "YTD", "All"]
PERF_CAP_PERIODS = ["1W", "1M", "3M", "6M", "1Y", "5Y", "YTD"]
RANGE_PERIODS = ["1M", "3M", "6M", "All"]
PLOT_SUFFIXES = [".upper", ".lower", ".middle", ".basis", ".basic", ".high", ".low"]

# Yatay hareketli ortalamalar: uzunluk anahtarın birliğinde (SMA20)
MA_LENGTHS = [10, 20, 50, 100, 200]
MA_FAMILIES = ["SMA", "EMA", "WMA", "VWMA", "HullMA", "DEMA", "TEMA", "SMMA", "ALMA", "LSMA"]

# Uzunluk parametresi olan göstergeler: alt çizgiyle (ADX_9, Mom_14)
LENGTH_PARAM_FAMILIES = [
    "ADX", "ADX+DI", "ADX-DI", "Mom", "BBPower", "Stoch.K", "Stoch.D",
    "Stoch.RSI.K", "Stoch.RSI.D", "Aroon.up", "Aroon.dn", "Donchian.middle",
    "Donchian.upper", "Donchian.lower", "KltChnl.middle", "KltChnl.upper",
    "KltChnl.lower", "W.R", "AROON", "AO", "UO", "CCI", "CMF", "MFI",
    "TSI", "ROC", "Ichimoku.BLine", "Stoch", "Aroon", "Ichimoku",
]

# Sabit isimli, parametresiz göstergeler
SIMPLE_INDICATORS = [
    "RSI", "ATR", "ATRP", "MACD.macd", "MACD.signal", "CCI20", "Mom", "BBPower",
    "AO", "AO[1]", "AO[2]", "UO", "W.R", "ADX", "ADX+DI", "ADX-DI", "Stoch.K",
    "Stoch.D", "Stoch.RSI.K", "Stoch.RSI.D", "P.SAR", "VWAP", "VWMA",
    "Ichimoku.BLine", "Aroon", "Donchian", "KltChnl", "BB.upper", "BB.lower",
    "BB.middle", "BBPower", "MoneyFlow.index", "ChaikinMoneyFlow", "Rec.Stoch.RSI",
    "Rec.WR", "Rec.MACD.D", "Rec.BBPower", "Rec.VWMA", "Rec.Ichimoku",
    "Recommend.All", "Recommend.MA", "Recommend.Other", "recommendation_mark",
    "Volatility.D", "Volatility.W", "Volatility.M",
    "Pivot.M.Classic.S1", "Pivot.M.Classic.S2", "Pivot.M.Classic.S3",
    "Pivot.M.Classic.R1", "Pivot.M.Classic.R2", "Pivot.M.Classic.R3",
    "Pivot.M.Fibonacci.S1", "Pivot.M.Fibonacci.R1", "Pivot.M.Camarilla.S1",
    "Pivot.M.Camarilla.R1", "Pivot.M.Woodie.S1", "Pivot.M.Woodie.R1",
    "Pivot.M.Demark.S1", "Pivot.M.Demark.R1",
]

# Her zaman ilk sırada durmasını istediğimiz sütunlar
CORE_COLUMNS = [
    "name", "description", "sector", "industry", "exchange", "type", "currency",
    "country", "isin", "close", "open", "high", "low", "change", "volume",
    "market_cap_basic", "price_52_week_high", "price_52_week_low",
    "relative_volume_10d_calc", "average_volume_30d_calc",
    "price_earnings_ttm", "price_earnings_forward", "price_book_fq",
    "dividends_yield", "beta_1_year",
    "earnings_release_date", "earnings_release_next_date",
    "Recommend.All", "Recommend.MA", "Recommend.Other",
]

ALWAYS_LAST = [
    "name", "description", "sector", "industry", "exchange", "type", "currency",
    "country", "isin", "ticker-view", "active_symbol", "exchange_timezone",
    "fundamental_currency_code", "MarketValue", "MarketCap", "market_cap_to_tvl",
    "graham_numbers", "altman_z_score_fy", "altman_z_score_ttm",
    "piotroski_value", "pylonscore", "MagicFormula", "AltmanZScore",
    "analyst_rating", "AnalystRating", "MARating", "OsRating", "TechRating",
    "Value.Traded", "MoneyFlow", "Perf.All", "Perf.10Y", "Perf.3Y",
    "actively_managed", "private_company", "public_company", "esg_score",
]


# --------------------------------------------------------------------------
# yardımcılar
# --------------------------------------------------------------------------

def log(message: str) -> None:
    print(f"[TV] {message}", flush=True)


def _http(url: str, data: bytes | None = None, headers: dict | None = None,
          timeout: int = TIMEOUT, retries: int = RETRIES) -> bytes:
    last: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            request = urllib.request.Request(url, data=data, headers=headers or HEADERS)
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.read()
        except Exception as exc:
            last = exc
            if attempt < retries:
                time.sleep(1.2 * attempt)
    raise last  # type: ignore[misc]


def fetch_text(url: str, timeout: int = TIMEOUT) -> str:
    return _http(url, timeout=timeout).decode("utf-8", "replace")


def scan(market: str, columns: list[str], row_range: tuple[int, int] = (0, 1000),
         extra: dict | None = None) -> dict:
    body: dict = {
        "symbols": {"query": {"types": []}, "tickers": []},
        "columns": list(columns),
        "range": [row_range[0], row_range[1]],
        "options": {"lang": LANG},
    }
    if extra:
        body.update(extra)
    payload = json.dumps(body).encode("utf-8")
    headers = dict(HEADERS)
    headers["Content-Type"] = "application/json"
    raw = _http(SCANNER_URL.format(market=market), data=payload, headers=headers)
    return json.loads(raw.decode("utf-8"))


def scalarize(value):
    """Bazi sutunlar nesne donduruyor; Excel'e yazilabilir tek degere indirger."""
    if isinstance(value, dict):
        for field in ("close", "value", "name", "description", "text", "title"):
            inner = value.get(field)
            if isinstance(inner, (int, float, str)) and not isinstance(inner, bool):
                return inner
        parts = [f"{key}={item}" for key, item in value.items()
                 if not isinstance(item, (dict, list, tuple))]
        return ", ".join(parts) if parts else None
    if isinstance(value, (list, tuple)):
        return ", ".join(str(item) for item in value) if value else None
    return value


# --------------------------------------------------------------------------
# 1) Sütun kataloğunun keşfi (frontend paketlerinden)
# --------------------------------------------------------------------------

BUNDLE_RE = re.compile(
    r"https://static\.tradingview\.com/static/bundles/[A-Za-z0-9_\-\.]+\.js"
)
CLASS_RE = re.compile(
    r'_getColumnKey\(\)\{(?P<body>.*?)\}constructor\('
    r'(?P<ctor>[^;]*?this\._key="(?P<key>[^"]+)")',
    re.S,
)


def find_column_bundles(max_misses: int = 12) -> list[str]:
    html = fetch_text(SCREENER_PAGE)
    urls = list(dict.fromkeys(BUNDLE_RE.findall(html)))
    log(f"screener sayfasında {len(urls)} JS paketi bulundu, sütun tanımları taranıyor...")
    hits: list[str] = []
    misses = 0
    for url in reversed(urls):
        try:
            text = fetch_text(url, timeout=45)
        except Exception as exc:
            log(f"  ! paket atlandı: {url.rsplit('/', 1)[-1]} ({exc})")
            continue
        if '_key="' in text and "AbstractValueTableColumn" in text:
            log(f"  + {url.rsplit('/', 1)[-1]}")
            hits.append(url)
        else:
            misses += 1
            if hits and misses >= max_misses:
                break
    if not hits:
        raise RuntimeError("Sütun tanımları bulunamadı; sayfa yapısı değişmiş olabilir.")
    return hits


def extract_column_keys(bundle_urls: list[str]) -> tuple[list[str], dict[str, set]]:
    keys: set[str] = set()
    sort_keys: set[str] = set()
    params: dict[str, set] = {}
    for url in bundle_urls:
        text = fetch_text(url, timeout=60)
        keys.update(re.findall(r'_key="([^"]+)"', text))
        sort_keys.update(re.findall(r'sortKey="([^"]+)"', text))
        for match in CLASS_RE.finditer(text):
            key = match.group("key")
            body = match.group("body")
            info = params.setdefault(key, set())
            if "FiscalPeriod" in body:
                info.add("fiscal")
            if "Period." in body:
                info.add("period")
            if "Length" in body:
                info.add("length")
            if "Interval" in body:
                info.add("interval")
            if "Plot" in body:
                info.add("plot")
            if "_calc" in body:
                info.add("calc")
    return sorted(keys | sort_keys), params


def build_candidates(base_keys: list[str], params: dict[str, set]) -> list[str]:
    candidates: list[str] = list(CORE_COLUMNS)

    def add(value: str) -> None:
        if value:
            candidates.append(value)

    for key in base_keys:
        if key in ALWAYS_LAST:
            continue
        add(key)
        info = params.get(key, set())
        if "fiscal" in info:
            for suffix in FISCAL_SUFFIXES:
                add(key + suffix)
        if "period" in info:
            for suffix in GROWTH_SUFFIXES:
                add(key + suffix)
        if "length" in info:
            for length in LENGTHS:
                add(f"{key}_{length}")
        if "interval" in info:
            for length in CALC_LENGTHS:
                add(f"{key}_{length}_calc")
        if "plot" in info:
            for plot in PLOT_SUFFIXES:
                add(key + plot)

    # zaman dilimli performans
    for period in PERF_PERIODS:
        add(f"Perf.{period}")
    for period in PERF_CAP_PERIODS:
        add(f"Perf.{period}.MarketCap")
    for suffix in CHANGE_SUFFIXES:
        add("change" + suffix)
    for period in RANGE_PERIODS:
        add(f"High.{period}")
        add(f"Low.{period}")

    # hareketli ortalamalar
    for family in MA_FAMILIES:
        for length in MA_LENGTHS:
            add(f"{family}{length}")

    # uzunluk parametreli göstergeler
    for family in LENGTH_PARAM_FAMILIES:
        for length in LENGTHS:
            add(f"{family}_{length}")

    # göstergeler ve zaman dilimleri
    for indicator in SIMPLE_INDICATORS:
        add(indicator)
        for tf in TF_SUFFIXES:
            add(indicator + tf)
    for family in MA_FAMILIES:
        for tf in TF_SUFFIXES:
            add(f"{family}20{tf}")

    unique = list(dict.fromkeys(candidates))
    log(f"toplam aday sütun: {len(unique)}")
    return unique


def load_catalog(force_refresh: bool = False) -> tuple[list[str], dict]:
    if not force_refresh and CATALOG_FILE.exists():
        try:
            data = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))
            candidates = data.get("candidates") or []
            if candidates:
                log(f"sütun kataloğu önbellekten: {len(candidates)} aday "
                    f"({data.get('built_at', '?')})")
                return candidates, data
        except Exception as exc:
            log(f"katalog önbelleği okunamadı ({exc}), yeniden keşfedilecek")

    log("sütun kataloğu yeniden keşfediliyor (sitenin 'Sütun ekle' listesi)...")
    bundles = find_column_bundles()
    base_keys, params = extract_column_keys(bundles)
    log(f"paketlerden {len(base_keys)} temel sütun anahtarı çıkarıldı")
    candidates = build_candidates(base_keys, params)
    meta = {
        "built_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "bundles": [url.rsplit("/", 1)[-1] for url in bundles],
        "base_key_count": len(base_keys),
        "candidate_count": len(candidates),
        "candidates": candidates,
    }
    try:
        CATALOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CATALOG_FILE.write_text(
            json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
        log(f"katalog kaydedildi: {CATALOG_FILE}")
    except Exception as exc:
        log(f"katalog yazılamadı ({exc})")
    return candidates, meta


# --------------------------------------------------------------------------
# 2) Veri çekme
# --------------------------------------------------------------------------

def fetch_all(market: str, columns: list[str]) -> tuple[list[str], dict]:
    probe = scan(market, ["name"], (0, 1000))
    symbols = [row["s"] for row in probe.get("data") or []]
    total = probe.get("totalCount", len(symbols))
    if not symbols:
        raise RuntimeError("Hisse listesi boş geldi.")
    log(f"{market}: {len(symbols)}/{total} sembol, {len(columns)} sütun için veri çekiliyor...")

    index = {symbol: position for position, symbol in enumerate(symbols)}
    data: dict[str, list] = {column: [None] * len(symbols) for column in columns}
    batches = 0
    for start in range(0, len(columns), BATCH):
        chunk = columns[start:start + BATCH]
        result = scan(market, chunk, (0, 1000))
        for row in result.get("data") or []:
            position = index.get(row["s"])
            if position is None:
                continue
            values = row["d"]
            for offset, column in enumerate(chunk):
                if offset < len(values):
                    data[column][position] = values[offset]
        batches += 1
        done = min(start + BATCH, len(columns))
        print(f"\r       ilerleme: {done}/{len(columns)} sütun ({batches} istek)", end="", flush=True)
        time.sleep(0.1)
    print()

    dirty = [column for column in columns
             if any(isinstance(value, (dict, list, tuple)) for value in data[column])]
    if dirty:
        log(f"{len(dirty)} sütun nesne döndürüyor, metne çevriliyor: {', '.join(dirty[:8])}")
        for column in dirty:
            data[column] = [scalarize(value) for value in data[column]]
    return symbols, data


# --------------------------------------------------------------------------
# 3) Turkce etiketler
# --------------------------------------------------------------------------

PERIOD_LABELS = {
    "1": "1 G\u00fcn", "2": "2 G\u00fcn", "3": "3 G\u00fcn", "5": "5 G\u00fcn",
    "10": "10 G\u00fcn", "15": "15 G\u00fcn", "30": "1 Ay", "45": "1,5 Ay",
    "60": "2 Ay", "90": "3 Ay", "120": "4 Ay", "180": "6 Ay", "240": "1 Y\u0131l",
    "1W": "1 Hafta", "2W": "2 Hafta", "1M": "1 Ay", "2M": "2 Ay", "3M": "3 Ay",
    "6M": "6 Ay", "YTD": "Y\u0131l Ba\u015f\u0131", "1Y": "1 Y\u0131l", "2Y": "2 Y\u0131l",
    "3Y": "3 Y\u0131l", "5Y": "5 Y\u0131l", "10Y": "10 Y\u0131l", "20Y": "20 Y\u0131l",
    "All": "T\u00fcm Zamanlar", "W": "Haftal\u0131k", "52W": "52 Hafta",
    "250D": "250 G\u00fcn", "250": "250 G\u00fcn", "180": "6 Ay",
    "D": "G\u00fcnl\u00fck", "1D": "G\u00fcnl\u00fck",
}

TF_LABELS = {
    "1": "1 dk", "5": "5 dk", "15": "15 dk", "30": "30 dk", "60": "1 saat",
    "120": "2 saat", "240": "4 saat", "1W": "Haftal\u0131k", "1M": "Ayl\u0131k",
}

TOKEN_LABELS = {
    "total": "Toplam", "net": "Net", "gross": "Br\u00fct", "operating": "Faaliyet",
    "revenue": "Gelir", "sales": "Sat\u0131\u015f", "income": "Gelir",
    "profit": "K\u00e2r", "earnings": "Kazanc", "assets": "Varl\u0131k",
    "liabilities": "Y\u00fck\u00fcm\u00fcl\u00fck", "equity": "\u00d6zkaynak",
    "cash": "Nakit", "debt": "Bor\u00e7", "margin": "Marj", "growth": "B\u00fcy\u00fcyme",
    "yield": "Verim", "ratio": "Oran", "share": "Hisse", "per": "/",
    "to": "/", "vs": " / ",
    "basic": "Baz", "diluted": "Seyreltilmi\u015f", "payout": "Da\u011f\u0131t\u0131m",
    "dividend": "Temett\u00fc", "dividends": "Temett\u00fc", "book": "Defter",
    "value": "De\u011fer", "market": "Piyasa", "cap": "De\u011feri",
    "high": "En Y\u00fcksek", "low": "En D\u00fc\u015f\u00fck", "open": "A\u00e7\u0131l\u0131\u015f",
    "close": "Kapan\u0131\u015f", "avg": "Ort.", "average": "Ortalama",
    "volume": "Hacim", "change": "De\u011fi\u015fim", "performance": "Performans",
    "return": "Getiri", "employee": "\u00c7al\u0131\u015fan", "employees": "\u00c7al\u0131\u015fan", "rnd": "Ar-Ge",
    "research": "Ar-Ge", "dev": "Geli\u015ftirme", "estimate": "Tahmin",
    "forecast": "Tahmin", "actual": "Ger\u00e7ekle\u015fen", "annual": "Y\u0131ll\u0131k",
    "tax": "Vergi", "interest": "Faiz", "expense": "Gider", "costs": "Maliyet",
    "flow": "Ak\u0131\u015f", "free": "Serbest", "capital": "Sermaye",
    "expenditure": "Harcama", "investment": "Yat\u0131r\u0131m",
    "outstanding": "Dola\u015f\u0131mda", "shares": "Hisse", "float": "Halka A\u00e7\u0131k",
    "volatility": "Volatilite", "relative": "G\u00f6reli", "typical": "Tipik",
    "target": "Hedef", "analyst": "Analist", "sector": "Sekt\u00f6r",
    "industry": "Alt Sekt\u00f6r", "exchange": "Borsa", "currency": "Para Birimi",
    "number": "Say\u0131", "of": "", "and": "ve", "with": "",
    "recommend": "Tavsiye", "all": "T\u00fcm", "ma": "Hareketli Ort.",
    "other": "Di\u00efer", "percent": "Y\u00fczde", "days": "G\u00fcn",
    "day": "G\u00fcn", "month": "Ay", "year": "Y\u0131l", "quarter": "\u00c7eyrek",
    "fy": "Y\u0131ll\u0131k", "fq": "\u00c7eyreklik", "ttm": "TTM",
    "yoy": "YOY", "qoq": "\u00c7eyreklik", "cagr": "Bile\u015fik Y\u0131ll\u0131k B\u00fcy\u00fcyme",
    "scores": "Skor", "score": "Skor", "calculation": "Hesaplama",
    # piyasa / tahvil / varlık kiralama terimleri
    "price": "Fiyat", "fwd": "İleri", "ask": "Alış", "bid": "Satış",
    "bond": "Tahvil", "bonds": "Tahviller", "coupon": "Kupon",
    "maturity": "Vade", "redemption": "Anapara Ödemesi", "issuer": "İhraççı",
    "next": "Sonraki", "prev": "Önceki", "general": "Genel", "gen": "Genel",
    "after": "Sonrası", "long": "Uzun", "short": "Kısa",
    "liquidations": "Tasfiye", "estimate": "Tahmin", "provision": "Karşılık",
    "traded": "İşlem", "avg": "Ortalama", "offer": "Teklif",
    "date": "Tarih", "admin": "İdari", "exp": "Gider", "sales": "Satış",
    "activities": "Faaliyetler", "position": "Pozisyon", "strength": "Güç",
    "risk": "Risk", "rank": "Sıra", "count": "Sayı", "issue": "İhraç",
    "parent": "Ana", "outlook": "Görünüm", "call": "Call", "put": "Put",
    "medical": "T\u0131bbi", "services": "Hizmetler", "products": "\u00dcr\u00fcnler",
    "technology": "Teknoloji", "goods": "Malzeme", "trading": "Ticaret",
    "exploration": "Arama", "production": "\u00dcretim", "development": "Geli\u015ftirme",
    "revenuegrowth": "", "sharesoutstanding": "", "tomorrow": "",
}

PLOT_LABELS = {
    "upper": "\u00dcst", "lower": "Alt", "middle": "Orta", "basis": "Taban",
    "basic": "Baz", "high": "Y\u00fcksek", "low": "D\u00fc\u015f\u00fck",
    "s1": "S1", "s2": "S2", "s3": "S3", "r1": "R1", "r2": "R2", "r3": "R3",
}

# teknik gösterge kökenleri: teknik birliğinde olduğu bilinen anahtarlar
TECHNICAL_ROOTS = {
    "ADX", "ADX+DI", "ADX-DI", "ADR", "ADRP", "AO", "AO[1]", "AO[2]", "ATR", "ATRP",
    "BBPower", "CCI20", "MACD.macd", "MACD.signal", "Mom", "RSI", "Stoch.K", "Stoch.D",
    "Stoch.RSI.K", "Stoch.RSI.D", "UO", "W.R", "P.SAR", "VWAP", "VWMA", "Aroon",
    "Aroon.up", "Aroon.dn", "Donchian", "Donchian.middle", "Ichimoku",
    "Ichimoku.BLine", "KltChnl", "KltChnl.middle", "BB.upper", "BB.lower",
    "BB.middle", "ChaikinMoneyFlow", "MoneyFlow.index", "Rec.Stoch.RSI", "Rec.WR",
    "Rec.MACD.D", "Rec.BBPower", "Rec.VWMA", "Rec.Ichimoku", "Volatility.D",
    "Volatility.W", "Volatility.M", "Pivot.M.Classic", "Pivot.M.Fibonacci",
    "Pivot.M.Camarilla", "Pivot.M.Woodie", "Pivot.M.Demark", "SMA", "EMA", "WMA",
    "HullMA", "DEMA", "TEMA", "SMMA", "ALMA", "LSMA", "Stoch", "Stoch.RSI", "MACD",
    "AROON", "BB", "ROC", "CMF", "MFI", "TSI",
}
TECHNICAL_PREFIXES = ("SMA", "EMA", "WMA", "HullMA", "DEMA", "TEMA", "ALMA",
                      "SMMA", "LSMA", "VWMA", "ADX", "AO[", "CCI", "ATR",
                      "BBPower", "Mom", "RSI", "Stoch", "AROON", "Rec.")
FUND_HINTS = (
    "aum", "etf", "fund", "brand", "cfi", "holdings", "issuer", "k1_form", "ucits",
    "weight_top", "weighting_scheme", "transparent_holding", "leveraged", "leverage",
    "niche", "focus", "strategy", "selection_criteria", "ipo_", "is_primary",
    "net_asset", "nav_", "holds_derivatives", "launch_date", "category",
)
PRICE_HINTS = (
    "close", "open", "high", "low", "price", "volume", "gap", "nav", "vwap",
    "vma", "bid", "ask", "turnover", "drawdown", "relative_volume",
)
EXTRA_FINANCIAL_HINTS = (
    "dps_", "fiscal_period", "interst_cover", "zmijewski", "score", "cover",
    "last_price_update", "fiscal_year", "tax_rate",
)

CATEGORY_RULES = [
    (("Recommend", "recommendation_mark", "MARating", "OsRating", "TechRating",
      "AnalystRating", "analyst_rating"), "Teknik \u00b7 Tavsiye"),
    (("Perf.", "Volatility.", "change", "High.", "Low.", "price_52_week",
      "return_over", "alpha", "sortino", "sharpe"), "Performans"),
    (("price_earnings", "price_book", "price_sales", "price_to_", "enterprise_value",
      "market_cap", "MarketCap", "ev_", "peg", "beta", "payout_ratio",
      "dividends_yield", "yield", "graham", "altman", "piotroski", "pylonscore",
      "MagicFormula", "fcf_yield", "fcf_yield_recent"), "De\u011ferleme"),
    (("dividend", "payout", "buyback", "continuous_dividend"), "Temett\u00fc"),
    (("yoy_growth", "qoq_growth", "cagr_", "earnings_per_share_forecast",
      "eps", "revenue_growth", "income_growth", "growth"), "Finansal \u00b7 B\u00fcy\u00fcyme"),
    (("earnings_release", "forecast", "estimate", "target_price", "analyst",
      "number_of_analysts", "recommendation"), "Analist"),
    (("employee", "per_employee", "insider", "institutional", "major_holders",
      "shares_outstanding", "float", "held_by", "owned_by", "holder",
      "company_profile", "employees", "ceo", "founded", "headquarters",
      "website", "phone", "fax", "address"), "Şirket"),
    (("sector", "industry", "exchange", "currency", "type", "name",
      "description", "country", "isin", "timezone", "index", "primary_listing",
      "listing", "lot_size", "market", "symbol", "ticker"), "Kimlik"),
]


def _translate_token(token: str) -> str:
    lower = token.lower()
    if lower in TOKEN_LABELS:
        return TOKEN_LABELS[lower]
    if lower in PLOT_LABELS:
        return PLOT_LABELS[lower]
    if lower.isdigit():
        return token
    if re.fullmatch(r"[A-Za-z0-9]{2,12}", token):
        if any(character.isupper() for character in token[1:]):
            return token
        return token.capitalize()
    return token


def humanize(key: str) -> str:
    """Anahtari okunabilir Turkce etikete cevirir."""
    if not key:
        return key

    # 1) zaman dilimi son eki: "...|1" -> "... (1 dk)"
    timeframe = ""
    body = key
    if "|" in body:
        head, _, tail = body.rpartition("|")
        if re.fullmatch(r"[0-9A-Za-z]+", tail):
            timeframe = TF_LABELS.get(tail, tail)
            body = head
    # 2) Perf.<dönem>[.MarketCap]
    perf_tail = ""
    match = re.match(r"^Perf\.([0-9A-Za-z]+)(?:\.MarketCap)?$", body)
    if match:
        period = PERIOD_LABELS.get(match.group(1), match.group(1))
        is_cap = body.endswith(".MarketCap")
        base = "Piyasa De\u011feri Performans\u0131" if is_cap else "Performans"
        return f"{base} ({period})"
    # 3) High.<dönem> / Low.<dönem>
    match = re.match(r"^(High|Low)\.([0-9A-Za-z]+)$", body)
    if match:
        period = PERIOD_LABELS.get(match.group(2), match.group(2))
        word = "En Y\u00fcksek" if match.group(1) == "High" else "En D\u00fc\u015f\u00fck"
        return f"{period} {word} Fiyat"
    # 4) change|<dönem>
    match = re.match(r"^change\|([0-9A-Za-z]+)$", key)
    if match:
        period = PERIOD_LABELS.get(match.group(1), match.group(1))
        return f"De\u011fi\u015fim ({period})"
    # 5) büyüme sonekleri
    for suffix, label in (
        ("_yoy_growth_ttm", "B\u00fcy\u00fcyme (TTM, YOY)"),
        ("_yoy_growth_fy", "B\u00fcy\u00fcyme (Y\u0131ll\u0131k, YOY)"),
        ("_yoy_growth_fq", "B\u00fcy\u00fcyme (\u00c7eyreklik, YOY)"),
        ("_qoq_growth_fq", "B\u00fcy\u00fcyme (\u00c7eyreklik, QoQ)"),
        ("_cagr_5y", "5 Y\u0131ll\u0131k Bile\u015fik B\u00fcy\u00fcyme"),
    ):
        if body.endswith(suffix):
            body = body[: -len(suffix)]
            tail_label = label
            break
    else:
        tail_label = ""
        for suffix, label in (("_ttm", "(TTM)"), ("_fq", "(\u00c7eyreklik)"),
                              ("_fy", "(Y\u0131ll\u0131k)")):
            if body.endswith(suffix):
                body = body[: -len(suffix)]
                tail_label = label
                break
    # 6) _30d_calc gibi hacim ortalamalari
    match = re.search(r"_(\d+)(?:d)?_calc$", body)
    if match:
        days = match.group(1)
        body = body[: match.start()]
        tail_label = f"({PERIOD_LABELS.get(days, days)} Ortalama)"
    # 7) goreceli hacim
    body = re.sub(r"^relative_volume_(\d+)d$",
                  "G\u00f6reli Hacim (\\1 g\u00fcn)", body)

    words = [_translate_token(word)
             for word in re.split(r"[._|]", body) if word]
    words = [word for word in words if word]
    label = re.sub(r"\s+", " ", " ".join(words)).strip()
    if tail_label:
        label = f"{label} {tail_label}".strip()
    if timeframe:
        label = f"{label} ({timeframe})" if label else timeframe
    return label or key


def make_labeler(overrides: dict) -> callable:
    cache: dict[str, str] = {}

    def label_for(key: str) -> str:
        if key in cache:
            return cache[key]
        result = overrides.get(key) or humanize(key)
        if len(result) > 58:
            result = result[:55] + "..."
        cache[key] = result
        return result

    return label_for


def make_unique_labeler(label_for, columns: list[str]) -> callable:
    """Ayni Turkce etiket birden fazla sutunda tekrar ederse ayirt edici ek ekler."""
    seen: dict[str, int] = {}
    result: dict[str, str] = {}
    for key in columns:
        label = label_for(key)
        count = seen.get(label, 0)
        if count:
            # "toplam_revenue_fy" / "_fq" gibi durumlarda son eki goster
            match = re.search(r"_(fy|fq|ttm|annual|quarter|1y|3y|5y|10y)$", key)
            if match:
                suffix_label = {"fy": "(Y\u0131ll\u0131k)", "fq": "(\u00c7eyreklik)",
                                "ttm": "(TTM)", "annual": "(Y\u0131ll\u0131k)",
                                "quarter": "(\u00c7eyreklik)", "1y": "(1 Y\u0131l)",
                                "3y": "(3 Y\u0131l)", "5y": "(5 Y\u0131l)",
                                "10y": "(10 Y\u0131l)"}[match.group(1)]
                label = f"{label} {suffix_label}"
            else:
                tail = key.rsplit("_", 1)[-1] if "_" in key else key
                label = f"{label} [{tail}]"
            base, attempt = label, 2
            while label in seen and seen[label] > 0:
                label = f"{base} ({attempt})"
                attempt += 1
        seen[label] = seen.get(label, 0) + 1
        result[key] = label
    return lambda key: result.get(key, label_for(key))


FINANCIAL_HINTS = (
    "revenue", "income", "profit", "ebit", "asset", "liabilit", "equity", "cash",
    "debt", "tax", "expense", "cost", "capital", "share", "dividend", "book_value",
    "receivable", "inventor", "payable", "goodwill", "investment", "working_capital",
    "research", "rnd", "opex", "capex", "depreciation", "amortization", "interest",
    "loan", "borrowing", "deposit", "nopat", "margin", "ratio", "per_employee",
    "yield", "payout", "coverage", "turnover", "growth", "salary", "employee",
    "intangible", "deferred", "minority", "preferred", "lease", "pension",
    "provision", "impairment", "expenses", "sga", "rental", "royalty", "revenuegrowth",
)


def is_technical(key: str) -> bool:
    if key in TECHNICAL_ROOTS or key.startswith(TECHNICAL_PREFIXES):
        return True
    for root in sorted(TECHNICAL_ROOTS, key=len, reverse=True):
        if key.startswith((root + ".", root + "|", root + "_", root + "(")):
            return True
    return bool(re.fullmatch(r"[A-Z][A-Za-z0-9.\[\]]{0,8}\d{1,3}(?:_calc)?", key))


def category_of(key: str) -> str:
    for needles, category in CATEGORY_RULES:
        if any(needle in key for needle in needles):
            return category
    if is_technical(key):
        return "Teknik"
    lower = key.lower()
    if any(hint in lower for hint in FUND_HINTS):
        return "Fon / ETF"
    if any(hint in lower for hint in PRICE_HINTS):
        return "Fiyat \u00b7 Hacim"
    if any(hint in lower for hint in FINANCIAL_HINTS + EXTRA_FINANCIAL_HINTS):
        return "Finansal"
    return "Di\u011fer"


# --------------------------------------------------------------------------
# 4) Sayı biçimleri
#
# TradingView yüzde sütunlarını 12.62 gibi döndürüyor (0.1262 değil),
# bu yüzden Excel'in yüzde biçimi 100 ile çarpılmamalı: 0.00"%" kullanılır.
# --------------------------------------------------------------------------

PERCENT_EXACT = {
    "change", "gross_margin", "net_margin", "operating_margin", "pre_tax_margin",
    "profit_margin", "return_on_equity", "return_on_assets",
    "return_on_invested_capital", "dividends_yield", "payout_ratio",
    "recommendation_mark", "efficiency", "fiscal_year",
}
PERCENT_SUBSTRINGS = (
    "margin", "yoy_growth", "qoq_growth", "cagr_", "dividend_yield",
    "dividends_yield", "payout_ratio", "return_on_equity", "return_on_assets",
    "return_on_invested", "Perf.", "Volatility.", "change|", "High.", "Low.",
    "price_52_week", "yield", "inflation", "premium", "discount", "chance",
    "rate", "tax", "roi", "roe", "roa", "roic", "growth",
)
MONEY_SUBSTRINGS = (
    "market_cap", "volume", "_calc", "revenue", "income", "profit", "cash",
    "ebitda", "ebit", "assets", "liabilities", "equity", "debt", "capex",
    "expenditure", "shares", "employees", "dividends", "free_cash",
    "operating_cash", "book_value", "enterprise_value", "invested_capital",
    "net_debt", "float_shares", "working_capital", "gross_profit",
    "research_and_dev", "rd_", "rnd", "capex",
)
PRICE_KEYS = {"close", "open", "high", "low", "avg", "average_volume"}
# oran / katsayı türü sütunlar: para tutarı değil, ondalıklı karşılaştırma
RATIO_HINTS = (
    "_ratio", "ratio_", "to_equity", "to_total", "to_asset", "to_book", "to_sales",
    "to_cash", "to_revenue", "to_capital", "per_", "_per", "times", "multiple",
    "coverage", "debt_to", "quick_", "current_", "beta", "price_to", "price_book",
    "price_sales", "price_earnings", "peg", "book_value", "earnings_per",
    "recommendation_mark", "Recomm", "AnalystRating", "score", "pylonscore",
    "money_flow", "mfi", "relative_volume", "gap", "nav",
)


def number_format(key: str) -> str:
    if key in PERCENT_EXACT or any(item in key for item in PERCENT_SUBSTRINGS):
        return '0.00"%"'
    if key in PRICE_KEYS or key.startswith(("Pivot.M", "price_52_week",
                                           "High.", "Low.")):
        return "0.000"
    if is_technical(key):
        return "0.00"
    if any(item in key for item in RATIO_HINTS):
        return "0.00"
    if any(item in key for item in MONEY_SUBSTRINGS):
        return "#,##0"
    if "price" in key or "close" in key:
        return "0.000"
    return "General"


# --------------------------------------------------------------------------
# 5) Excel yazımı
# --------------------------------------------------------------------------

def export_excel(path: Path, sheet_name: str, symbols: list[str], columns: list[str],
                 data: dict, label_for, catalog_rows: list[dict]) -> None:
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = Workbook(write_only=True)

    head_font = Font(bold=True, color="FFFFFF", size=10)
    head_fill = PatternFill("solid", fgColor="1F3864")
    key_font = Font(bold=True, color="FFFFFF", size=8)
    key_fill = PatternFill("solid", fgColor="2F5597")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sheet = workbook.create_sheet(sheet_name[:31])
    widths = [14, 16]
    widths.extend([16] * len(columns))
    for position, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(position)].width = width
    # write_only modda sayfa ayarları ilk satır yazılmadan ÖNCE yapılmalıdır
    sheet.freeze_panes = "C3"

    header_row = []
    for text in ["BIST Kodu", "Hisse Adı"] + [label_for(column) for column in columns]:
        cell = WriteOnlyCell(sheet, value=text)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        header_row.append(cell)
    sheet.append(header_row)

    key_row = [WriteOnlyCell(sheet, value="symbol"), WriteOnlyCell(sheet, value="name")]
    for column in columns:
        cell = WriteOnlyCell(sheet, value=column)
        cell.font = key_font
        cell.fill = key_fill
        cell.alignment = center
        key_row.append(cell)
    sheet.append(key_row)

    formats = {column: number_format(column) for column in columns}
    name_values = data.get("name") or [None] * len(symbols)
    for position in range(len(symbols)):
        row = [symbols[position].split(":")[-1], name_values[position]]
        for column in columns:
            value = data[column][position]
            if value is None or isinstance(value, str):
                row.append(WriteOnlyCell(sheet, value=value))
            else:
                cell = WriteOnlyCell(sheet, value=value)
                cell.number_format = formats[column]
                row.append(cell)
        sheet.append(row)

    sheet.auto_filter.ref = f"A2:{get_column_letter(len(columns) + 2)}{len(symbols) + 2}"

    catalog_sheet = workbook.create_sheet("Sütun Kataloğu")
    catalog_sheet.freeze_panes = "A2"
    for letter, width in (("A", 7), ("B", 44), ("C", 54), ("D", 22), ("E", 14)):
        catalog_sheet.column_dimensions[letter].width = width
    head = []
    for text in ["Sıra", "TradingView Anahtarı", "Türkçe Etiket", "Kategori", "Dolu Veri"]:
        cell = WriteOnlyCell(catalog_sheet, value=text)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        head.append(cell)
    catalog_sheet.append(head)
    for item in catalog_rows:
        catalog_sheet.append([item["index"], item["key"], item["label"],
                              item["category"], item["filled"]])

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def write_catalog_csv(path: Path, columns: list[str], data: dict | None, label_for) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["sıra", "tradingview_anahtarı", "türkçe_etiket",
                         "kategori", "dolu_veri"])
        for index, column in enumerate(columns, start=1):
            filled = ""
            if data is not None:
                values = data.get(column) or []
                filled = "E" if any(value is not None for value in values) else "H"
            writer.writerow([index, column, label_for(column), category_of(column), filled])


def export_json(path: Path, sheet_name: str, symbols: list[str], columns: list[str],
                data: dict, label_for) -> None:
    """Tüm veriyi satır tabanlı JSON olarak yazar (her satır bir sembol)."""
    with path.open("w", encoding="utf-8") as handle:
        handle.write("[\n")
        last = len(symbols) - 1
        for position, symbol in enumerate(symbols):
            # Excel'deki "BIST Kodu" ile aynı düz kod + piyasa ön ekli tam ticker
            record = {"symbol": symbol.rsplit(":", 1)[-1], "ticker": symbol}
            for column in columns:
                record[column] = data[column][position]
            handle.write(json.dumps(record, ensure_ascii=False, default=str))
            handle.write(",\n" if position != last else "\n")
        handle.write("]\n")


# --------------------------------------------------------------------------
# 6) ana akış
# --------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        prog="tradingview_screener.py",
        description="TradingView tarama verilerini Excel'e aktarır.",
        epilog="Örnek: python tradingview_screener.py --sadece-hisse",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--market", default="turkey", metavar="KOD",
                        help="piyasa kodu: turkey, usa, germany, japan, global "
                             "(varsayılan: turkey)")
    parser.add_argument("--cikti", default="", metavar="DOSYA",
                        help="çıktı dosyasının yolu (varsayılan: ./Archive klasörü)")
    parser.add_argument("--katalog-yenile", action="store_true",
                        help="sütun kataloğunu yeniden keşfet (ön belleği yok sayar)")
    parser.add_argument("--katalog-yaz", default="", metavar="DOSYA",
                        help="sütun kataloğunu okunabilir CSV olarak da yaz")
    parser.add_argument("--sadece-hisse", action="store_true",
                        help="sadece hisseleri al (fon/ETF hariç)")
    parser.add_argument("--sutun", default="", metavar="ANAHTAR1,ANAHTAR2",
                        help="virgülle ayrılmış sütun anahtarları (boş = hepsi)")
    args = parser.parse_args()

    started = time.time()
    market = args.market
    sheet_name = MARKET_NAMES.get(market.lower(), market.upper()[:31])
    log(f"piyasa: {market} ({sheet_name})")

    candidates, _meta = load_catalog(force_refresh=args.katalog_yenile)
    if args.sutun:
        wanted = [item.strip() for item in args.sutun.split(",") if item.strip()]
        columns = [column for column in wanted if column in candidates]
        missing = [column for column in wanted if column not in candidates]
        if missing:
            log(f"katalogda bulunamayan sütunlar: {', '.join(missing)}")
    else:
        columns = candidates

    symbols, data = fetch_all(market, columns)

    if args.sadece_hisse:
        rows = scan(market, ["type"], (0, 1000)).get("data") or []
        order = {row["s"]: position for position, row in enumerate(rows)
                 if (row["d"] or [None])[0] == "stock"}
        pairs = [(index, order[symbol]) for index, symbol in enumerate(symbols)
                 if symbol in order]
        pairs.sort(key=lambda item: item[1])
        symbols = [symbol for symbol, _ in pairs]
        data = {column: [values[index] for index, _ in pairs]
                for column, values in data.items()}
        log(f"sadece hisse filtresi: {len(symbols)} sembol")

    live_columns = [column for column in columns
                    if any(value not in (None, "") for value in data[column])]
    dropped = len(columns) - len(live_columns)
    log(f"BIST verisi olan sütun: {len(live_columns)} "
        f"(veri olmadığı için elendi: {dropped})")
    # "name" sütunu Excel'de 2. kolon olarak ayrı yazılıyor
    live_columns = [column for column in live_columns if column != "name"]

    label_for = make_labeler(LABEL_MAP)
    label_for = make_unique_labeler(label_for, live_columns)
    sector_map = SECTOR_MAP
    for column in ("sector", "industry", "country"):
        if column in live_columns and sector_map:
            data[column] = [sector_map.get(value, value) if isinstance(value, str) else value
                            for value in data[column]]

    # önemli sütunleri başa al
    front = [column for column in CORE_COLUMNS if column in live_columns]
    front = [column for column in front if column not in ("name", "description")]
    rest = [column for column in live_columns if column not in front]
    live_columns = front + rest

    if args.katalog_yaz:
        target = Path(args.katalog_yaz)
        write_catalog_csv(target, live_columns, data, label_for)
        log(f"sütun kataloğu yazıldı: {target}")

    if args.cikti:
        target = Path(args.cikti)
    else:
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d")
        target = OUT_DIR / f"{sheet_name}_{stamp}.xlsx"
        # aynı gün birden çok çalıştırmada numara ekle, üzerine yazma
        counter = 2
        while target.exists():
            target = OUT_DIR / f"{sheet_name}_{stamp}_{counter}.xlsx"
            counter += 1

    catalog_rows = [
        {
            "index": index,
            "key": column,
            "label": label_for(column),
            "category": category_of(column),
            "filled": f"{sum(1 for v in data[column] if v is not None)} / {len(symbols)}",
        }
        for index, column in enumerate(live_columns, start=1)
    ]

    log(f"Excel yazılıyor: {target.name} ({len(symbols)} satır x {len(live_columns)} sütun)")
    export_excel(target, sheet_name, symbols, live_columns, data, label_for, catalog_rows)

    size_mb = target.stat().st_size / 1024 / 1024
    log(f"tamamlandı: {target} ({size_mb:.1f} MB, {time.time() - started:.1f} sn)")

    # Sabit isimli kopya + aynı isimli JSON (her çalıştırmada üzerine yazılır).
    # Yalnızca varsayılan çıktı klasörüne yazarken üretilir: --cikti ile
    # tek bir Excel istenmişse sırf on sütunluk bir deneme, 994 sütunluk
    # JSON'un üzerine yazıp veriyi bozmamalı.
    if target.parent == OUT_DIR:
        latest = OUT_DIR / "Bist_Hisse_Tum_Veriler.xlsx"
        shutil.copyfile(target, latest)
        log(f"kopya yazıldı: {latest.name} ({latest.stat().st_size / 1024 / 1024:.1f} MB)")

        json_target = OUT_DIR / "Bist_Hisse_Tum_Veriler.json"
        export_json(json_target, sheet_name, symbols, ["name"] + live_columns,
                    data, label_for)
        log(f"JSON yazıldı: {json_target.name} "
            f"({json_target.stat().st_size / 1024 / 1024:.1f} MB)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("durduruldu")
        sys.exit(130)
    except Exception as error:  # noqa: BLE001
        log(f"HATA: {error}")
        raise
