#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
農園 me! パッキングリスト 本番スクリプト
Misoca API から納品書を取得 → 集計 → B5 PDF 生成 → iCloud Drive に保存

使い方:
  python3 misoca_packing_main.py              # 当日の納品書
  python3 misoca_packing_main.py 2026-03-29   # 日付指定
"""

import json
import os
import re
import sys
from datetime import date, datetime

import requests
from reportlab.lib import colors
from reportlab.lib.pagesizes import B5
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ============================================================
# 設定
# ============================================================
TOKEN_FILE  = os.path.expanduser('~/.misoca_token.json')
ICLOUD_DIR  = os.path.expanduser(
    '~/Library/Mobile Documents/com~apple~CloudDocs/me-farm/COO/packing-list/output'
)
API_BASE    = 'https://app.misoca.jp/api/v3'
CLIENT_ID     = 'uIE3rh76mv5LqgJ6UjT5W8KDAgSaw86ruoHeX92IUC0'
CLIENT_SECRET = 'ICLx5meAFTYqB_5yNiW-a6lZQjQJmr_r7tDN4NSK9xE'
TOKEN_URL     = 'https://app.misoca.jp/oauth2/token'

TOKUSHU_CUSTOMERS = ['タケウチ', 'le Lotus', 'ロテュス', 'Le Lotus']

NOTE_MAP = {
    'レッドアマランサスs': '紙', 'レッドアマランサスM': '紙', 'レッドアマランサスL': '紙',
    'セロシア': '紙', 'オクラ': '紙', 'ハイビスカスローゼル': '紙',
    'マリーゴールドリーフ': '紙', 'マリーゴールドオレンジピール': '紙',
    'メキシカンマリーゴールド': '紙', 'ベゴニアグリーン': '紙', 'ベゴニアブロンズ': '紙',
    'つゆくさ': '紙', 'ジェノベーゼバジル': '紙', 'レモンバジル': '紙',
    'シナモンバジル': '紙', 'ダークオパールバジル': '紙', 'バジルミックス': '紙',
}

# ============================================================
# 商品マスター（GAS の MASTER_DATA_EMBEDDED と同一）
# ============================================================
MASTER_DATA = [
    {'genre': 'チルドレン', 'name': 'ハーブミックス', 'g': '100', 'pack': 'SP'},
    {'genre': 'チルドレン', 'name': 'ハーブミックス', 'g': '140', 'pack': 'SP'},
    {'genre': 'チルドレン', 'name': 'ハーブミックス', 'g': '150', 'pack': 'SP'},
    {'genre': 'チルドレン', 'name': 'ハーブミックス', 'g': '200', 'pack': 'SP'},
    {'genre': 'チルドレン', 'name': 'ハーブミックス', 'g': '250', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'チルドレン（100g）', 'g': '100', 'pack': 'SP'},
    {'genre': 'チルドレン', 'name': 'マイクロ➕花ミックス(ミニ)', 'g': '', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'チルドレン', 'g': '', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'with me!', 'g': '50', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'with me!', 'g': '70', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'with me!', 'g': '100', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'with me!', 'g': '140', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'with me!', 'g': '250', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドからし水菜', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドからし水菜', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ルッコラ', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ルッコラ', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'グリーンマスタード', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'グリーンマスタード', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドマスタード', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドマスタード', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'イタリアンレッド', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'イタリアンレッド', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ペッパークレス', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': '中葉春菊', 'g': '15', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': '中葉春菊(ミニ)', 'g': '5', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': '赤茎ダイコン', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ムラサキダイコン', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'カラフルダイコン', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ほうろくなたね', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ほうろくなたね', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダ水菜', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダ水菜', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'かぶ（桃色）', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'かぶ（桃色）', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'かぶ（紫色）', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'かぶ（紫色）', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダミックス', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダミックス', 'g': '100', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダミックス', 'g': '150', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダミックス', 'g': '225', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダミックス', 'g': '250', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ネギ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ネギ(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': '赤軸ほうれん草', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'アカザ', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ディル', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ディル(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'セルフィーユ', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'セルフィーユ(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'パスレー', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'パスレー(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'フェンネル', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'コリアンダー', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'アニス', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'にんじん', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'にんじん(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'セロリ', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'セロリ(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'みつば', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオゼイユs', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオゼイユM', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオゼイユM(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオゼイユL', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ハーブミックス', 'g': '7', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'テンドリルピー', 'g': '12', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'テンドリルピー', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ペリーラ青', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ペリーラ青(ミニ)', 'g': '3', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'ペリーラ赤', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'えごま', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レモンバーム', 'g': '4', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レモンバーム(ミニ)', 'g': '1.5', 'pack': 'ミニパック'},
    {'genre': 'マイクロリーフ', 'name': 'レッドアマランサスs', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドアマランサスM', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドアマランサスL', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'セロシア', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ひまわり', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'オクラ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ハイビスカスローゼル', 'g': '', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': '沖縄虹色ほうれん草', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'マリーゴールドリーフ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'マリーゴールドオレンジピール', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'メキシカンマリーゴールド', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ベゴニアグリーン', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ベゴニアブロンズ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'つゆくさ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'プルピエピンク', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'プルピエ斑入り', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'プルピエブラウンM', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'プルピエブラウンL', 'g': '20', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': '檸檬花椒菜 （レモンホォワジョーナ）', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ジェノベーゼバジル', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レモンバジル', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'シナモンバジル', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ダークオパールバジル', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'バジルミックス', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリス s', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリスM', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリスL', 'g': '4', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリスオータム', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリスオータムクラウン', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオキザリス s', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオキザリス M', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'レッドオキザリス L', 'g': '4', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'オキザリスネーブルライム', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'カモミール', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'クレイトニア', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'クレイトニア', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダバーネット', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'サラダバーネット', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'ヤロー', 'g': '5', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ヤロー', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'モフモフフェンネル', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ハコベ', 'g': '7', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'ハコベ', 'g': '15', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'スズメノエンドウ', 'g': '2', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'かきどおし', 'g': '25', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'かきどおし大', 'g': '25', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'かきどおし　斑入り', 'g': '25', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'かきどおし斑入り大', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'かきどおし斑入り花付き', 'g': '10', 'pack': '横SP'},
    {'genre': 'マイクロリーフ', 'name': 'オイスターリーフ', 'g': '10', 'pack': 'SP'},
    {'genre': 'マイクロリーフ', 'name': 'コルシカミント', 'g': '5', 'pack': 'SP'},
    {'genre': 'その他', 'name': '四つ葉のクローバー', 'g': '15', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ローズマリー', 'g': '20', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'タイム', 'g': '20', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'レモングラス', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'レモングラス', 'g': '100', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'レモンバーム大', 'g': '50', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'よもぎ', 'g': '50', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'アップルミント', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ストロベリーミント', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'モヒートミント', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': '和はっか', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ブロンズフェンネル', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'フレッシュハーブティーミックス', 'g': '100', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'セルバチコ', 'g': '100', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ナスタチウムリーフ s', 'g': '25', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ナスタチウムリーフ M', 'g': '25', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ナスタチウムリーフ M(ミニ)', 'g': '10', 'pack': 'ミニパック'},
    {'genre': 'その他', 'name': 'ナスタチウムリーフ L', 'g': '10', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'サンフラワーペタル オレンジ', 'g': '10', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'サンフラワーペタル レッド&イエロー', 'g': '10', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'サンフラワーペタル ブラウン&イエロー', 'g': '10', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'サンフラワーペタル ミックス', 'g': '10', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'マリーゴールド ペタル ミックス', 'g': '10', 'pack': '横SP'},
    {'genre': 'その他', 'name': 'ひまわりのつぼみ', 'g': '25', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナスタチウム', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオライエロー', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラオレンジ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラホワイト', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラブラック', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラロイヤルブルー', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラスカーレット', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラマリーナ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラライトローズ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラゴールデンブルース', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラブルーピコティ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラピンクハロー', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラパイナップルクラッシュ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラトリカラー', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラホァンホァン', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラバニーイヤーズ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラタイガーアイ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラアプリコットアンティーク', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラピンクアンティーク', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラプラムアンティーク', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラアンティークミックス', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ビオラミックス', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'マイクロビオラ', 'g': '22', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'フリルパンジー', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムホワイト', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムホワイト(ミニ)', 'g': '30', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムホワイト(ミニ)', 'g': '20', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムバイオレット', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムバイオレット(ミニ)', 'g': '30', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムディープローズ', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムディープローズ(ミニ)', 'g': '30', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムミックス', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'アリッサムミックス(ミニ)', 'g': '30', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'ストックホワイト', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ストックライトベージュ', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ストックピンク', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ストックローズ', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ストックバイオレット', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ストックミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'カレンデュライエロー', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'カレンデュラオレンジ', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'カレンデュラミックス', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーブルー', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーピンク', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーレッド', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーバイオレット', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーチョコレート', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワーミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナデシコレッド', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナデシコホワイト', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナデシコピンク', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナデシコローズ', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ナデシコミックス', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターレッド', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターホワイト', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターピンク', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターローズ', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターバイオレット', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ポップスターミックス', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ノースポール ホワイト', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ノースポール イエロー', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアイエロー', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアホワイト', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアピンク', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアスカーレット', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアバイオレット', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'リナリアミックス', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ボリジブルー', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ボリジホワイト', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ボリジミックス', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'トレニア ベイビーブルー', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'トレニア ベイビーピンク', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'トレニア ミックス', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドイエロー', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドオレンジ', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドサンライズ', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドレッドフォックス', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドストロベリーアンティーク', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドミックス', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールドミックス(ミニ)', 'g': '4', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'ミニマリーゴールド', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '黄花コスモス', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': '黄花コスモス(ミニ)', 'g': '4', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタスホワイト', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタスレッド', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタスピンク', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタスバイオレット', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタスミックス', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ベゴニアミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ベゴニアのつぼみ', 'g': '25', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニアレッド', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニアホワイト', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニアピンク', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニアミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きペチュニア', 'g': '12', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナレッド', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナホワイト', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナピンク', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナバイオレット', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナミックス', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バタフライピー', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'エディブルフラワーミックス', 'g': '', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'わさびルッコラ', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'オータムポエム', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'チーマディラーパ', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '白菜', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'オレンジはくさいの花（アンティークイエロー）', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'レッドからし水菜', 'g': '', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '水菜', 'g': '', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ブロッコリー', 'g': '', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ルッコラ', 'g': '20', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'ダイコン', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'セルバチコ', 'g': '10', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'さやだいこん', 'g': '25', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ほうろく菜花', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'のらぼう菜', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'エンドウ豆', 'g': '20', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '春菊', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'クレイトニア', 'g': '15', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'ヤロー', 'g': '10', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'ディル', 'g': '5', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'セルフィーユ', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コリアンダー', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'イタリアンパセリ', 'g': '5', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'みつば', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'セロリ', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'フェンネル', 'g': '5', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ブロンズフェンネル', 'g': '5', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'きゅうり', 'g': '5', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '檸檬花椒菜', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'ニラ', 'g': '10', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'チャイブ', 'g': '20', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'いちじくの葉っぱ10枚入り', 'g': '10', 'pack': 'SP'},
]

# ============================================================
# PDF デザイン定数
# ============================================================
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
FONT_NAME = 'HeiseiKakuGo-W5'

ROW_COLORS = [
    colors.HexColor('#EBF5FB'),
    colors.HexColor('#D6EAF8'),
    colors.HexColor('#EBF7ED'),
    colors.HexColor('#D5F5E3'),
    colors.HexColor('#FEF9E7'),
    colors.HexColor('#FDEBD0'),
]
UNKNOWN_COLOR  = colors.HexColor('#FFDAB9')
HEADER_BG      = colors.HexColor('#2C3E50')
HEADER_FG      = colors.white
HEADER_PACK_BG = colors.HexColor('#1E6B3C')
HEADER_PACK_FG = colors.HexColor('#D5F5E3')
TOTAL_BG       = colors.HexColor('#E8F8F5')

COL_HEADERS    = ['ジャンル', '品名', 'g数', '備考', 'SP', '横SP', 'MP', 'ミニ', 'タケウチ', 'ロテュス']
COL_WIDTHS_MM  = [24, 38, 12, 8, 11, 13, 11, 13, 13, 13]
COL_WIDTHS     = [w * mm for w in COL_WIDTHS_MM]


# ============================================================
# トークン管理
# ============================================================
def load_token() -> str:
    if not os.path.exists(TOKEN_FILE):
        print('❌ トークンファイルが見つかりません。misoca_auth.py を実行して認証してください。')
        sys.exit(1)
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    return data.get('access_token', '')


def refresh_token(refresh_tok: str) -> str:
    resp = requests.post(TOKEN_URL, data={
        'grant_type':    'refresh_token',
        'refresh_token': refresh_tok,
        'client_id':     CLIENT_ID,
        'client_secret': CLIENT_SECRET,
    })
    if resp.status_code != 200:
        return ''
    new_data = resp.json()
    with open(TOKEN_FILE, 'w') as f:
        json.dump(new_data, f, indent=2)
    return new_data.get('access_token', '')


def get_valid_token() -> str:
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    token = data.get('access_token', '')
    # 試し打ち
    r = requests.get(
        f'{API_BASE}/delivery_slips?per_page=1&page=1',
        headers={'Authorization': f'Bearer {token}'},
    )
    if r.status_code == 200:
        return token
    # 401 → リフレッシュ試行
    if r.status_code == 401 and data.get('refresh_token'):
        print('アクセストークン期限切れ。リフレッシュ中...')
        new_token = refresh_token(data['refresh_token'])
        if new_token:
            print('✅ トークンをリフレッシュしました。')
            return new_token
    print('❌ 認証エラー。misoca_auth.py を再実行して認証してください。')
    sys.exit(1)


# ============================================================
# Misoca API：納品書取得
# ============================================================
def fetch_delivery_slips(start_date: str, end_date: str) -> list:
    token = get_valid_token()
    headers = {'Authorization': f'Bearer {token}'}
    all_slips = []
    page = 1
    per_page = 20

    print(f'Misoca から納品書を取得中（{start_date} 〜 {end_date}）...')
    while page <= 50:
        url = (f'{API_BASE}/delivery_slips'
               f'?issue_date_from={start_date}&issue_date_to={end_date}'
               f'&per_page={per_page}&page={page}')
        r = requests.get(url, headers=headers)
        if r.status_code != 200:
            print(f'❌ API エラー: HTTP {r.status_code}')
            break
        data = r.json()
        if not isinstance(data, list) or len(data) == 0:
            break
        all_slips.extend(data)
        print(f'  ページ {page}: {len(data)} 件（累計: {len(all_slips)} 件）')
        if len(data) < per_page:
            break
        page += 1

    print(f'取得完了（フィルター前）: {len(all_slips)} 件')

    # APIの日付フィルターが不完全なため、クライアント側で発行日を再確認
    filtered = [s for s in all_slips if s.get('issue_date') == start_date]
    print(f'発行日 {start_date} のみ絞り込み後: {len(filtered)} 件')
    return filtered


# ============================================================
# データ処理（GAS の processAllSlips / aggregateItems と同一ロジック）
# ============================================================
def is_tokushu(customer_name: str) -> bool:
    return any(t in (customer_name or '') for t in TOKUSHU_CUSTOMERS)


KNOWN_GENRES = ['マイクロリーフ', 'エディブルフラワー', 'チルドレン', 'その他']

def parse_item_name(raw_name: str) -> dict:
    name = re.sub(r'\s+', ' ', (raw_name or '').strip())

    # ×N 形式のパック数を除去
    pack_count = 1
    m = re.search(r'[×x✕]\s*(\d+)', name, re.IGNORECASE)
    if m:
        pack_count = int(m.group(1)) or 1
        name = (name[:m.start()] + name[m.end():]).strip()

    # 先頭のジャンル名を検出・除去（スペースあり・なし両対応）
    detected_genre = ''
    for genre in KNOWN_GENRES:
        if name.startswith(genre + ' ') or name.startswith(genre + '\u3000'):
            detected_genre = genre
            name = name[len(genre):].strip()
            break
        elif name.startswith(genre) and len(name) > len(genre):
            detected_genre = genre
            name = name[len(genre):].strip()
            break

    # 末尾の数量表記を抽出して除去（7g / 25枚 / 22輪入り / 20本入り 等）
    qty_match = re.search(
        r'(\d+(?:\.\d+)?)\s*(?:g|輪入り?|本入り?|枚入り?|個入り?|枚|本|輪|個)$',
        name, re.IGNORECASE
    )
    g = ''
    if qty_match:
        g = qty_match.group(1)
        name = name[:qty_match.start()].strip()

    return {'baseName': name, 'g': g, 'packCount': pack_count, 'detectedGenre': detected_genre}


def find_pack_from_master(base_name: str, g_val: str = '', detected_genre: str = '') -> dict:
    """品名＋g値の両方が一致する場合のみヒット。どちらか不一致は要確認。"""
    def hit(m):
        return {'genre': m['genre'], 'pack': m['pack'],
                'master_name': m['name'], 'unknown': False}

    def norm(s):
        # スペース正規化＋全角括弧→半角括弧
        return re.sub(r'\s+', ' ', s.strip()).replace('（', '(').replace('）', ')')

    def nsp(s):
        return re.sub(r'\s+', '', norm(s))

    normalized = norm(base_name) if base_name else ''
    nospace    = nsp(base_name)  if base_name else ''

    # ① スペース・括弧正規化した品名 + g値で完全一致
    for m in MASTER_DATA:
        if norm(m['name']) == normalized and m['g'] == g_val:
            return hit(m)

    # ② スペース除去＋括弧正規化した品名 + g値で一致
    for m in MASTER_DATA:
        if nsp(m['name']) == nospace and m['g'] == g_val:
            return hit(m)

    # ③ ジャンル+品名の連結でマスター検索（例：エディブルフラワー+ミックス → エディブルフラワーミックス）
    if detected_genre and base_name:
        combined = nsp(detected_genre + base_name)
        for m in MASTER_DATA:
            if nsp(m['name']) == combined and m['g'] == g_val:
                return hit(m)

    # ④ 品名が空の場合、ジャンル名から始まるマスター項目を検索（例：チルドレン 100g）
    if detected_genre and not base_name:
        genre_norm = norm(detected_genre)
        for m in MASTER_DATA:
            if norm(m['name']).startswith(genre_norm) and m['g'] == g_val:
                return hit(m)

    # どれも合わず → 要確認
    return {'genre': '要確認', 'pack': 'SP', 'master_name': '', 'unknown': True}


def process_slips(slips: list) -> list:
    results = []
    for slip in slips:
        customer = slip.get('recipient_name') or slip.get('contact_name') or ''
        tokushu  = is_tokushu(customer)
        items    = slip.get('items') or slip.get('document_lines') or []
        for item in items:
            raw_name = item.get('name') or item.get('item_name') or ''
            # 送料・手数料など梱包不要の行は無視
            if re.search(r'送料|手数料|配送|運賃', raw_name):
                continue
            quantity = float(item.get('quantity') or item.get('count') or 1) or 1
            parsed   = parse_item_name(raw_name)
            master   = find_pack_from_master(parsed['baseName'], parsed['g'], parsed['detectedGenre'])

            if master['unknown']:
                # 要確認：Misocaの品名・g値をそのまま保持。ジャンルはプレフィックスから検出
                use_name  = parsed['baseName']
                use_g     = parsed['g']
                use_genre = parsed['detectedGenre'] or '要確認'
                use_pack  = 'SP'
            else:
                # マスター一致：マスターの正式名・g値・ジャンル・パックを使用
                use_name  = master['master_name']
                use_g     = parsed['g']   # ← Misocaのg値を集計に使う（変更しない）
                use_genre = master['genre']
                use_pack  = master['pack']

            # (ミニ) を含む品名は必ずミニパック列へ（マスター一致・未一致を問わず）
            if '(ミニ)' in use_name or '（ミニ）' in use_name:
                use_pack = 'ミニパック'

            results.append({
                'customerName': customer,
                'rawName':      raw_name,
                'baseName':     use_name,
                'g':            use_g,
                'quantity':     quantity,
                'packCount':    parsed['packCount'],
                'genre':        use_genre,
                'pack':         use_pack,
                'unknown':      master['unknown'],
                'isTokushu':    tokushu,
            })
    return results


def aggregate_items(items: list) -> list:
    agg = {}
    for item in items:
        key = item['genre'] + '|' + item['baseName'] + '|' + item['g']
        if key not in agg:
            agg[key] = {
                'genre': item['genre'], 'baseName': item['baseName'],
                'g': item['g'], 'pack': item['pack'], 'unknown': item['unknown'],
                'sp': 0, 'yokoSP': 0, 'mp': 0, 'mini': 0,
                'takeuchi': 0, 'lotus': 0,
            }
        qty   = item['quantity'] * item['packCount']
        g_num = float(item['g']) if item['g'] else 0
        cn    = item['customerName']

        is_takeuchi = 'タケウチ' in cn
        is_lotus    = any(t in cn for t in ['le Lotus', 'Le Lotus', 'ロテュス'])

        if item['genre'] == 'マイクロリーフ' and (is_takeuchi or is_lotus):
            # マイクロリーフのみ：g×数量をタケウチ・ロテュス列へ
            if is_takeuchi:
                agg[key]['takeuchi'] += g_num * qty
            else:
                agg[key]['lotus'] += g_num * qty
        else:
            # エディブルフラワー・その他（タケウチ・ロテュス含む）→ 通常のパック列に合算
            p = item['pack']
            if p == 'SP':           agg[key]['sp']     += qty
            elif p == '横SP':       agg[key]['yokoSP'] += qty
            elif p == 'MP':         agg[key]['mp']     += qty
            elif p == 'ミニパック': agg[key]['mini']   += qty
            else:                   agg[key]['sp']     += qty
    return list(agg.values())


def build_rows_in_master_order(aggregated: list) -> list:
    """マスター順に並べ、注文ゼロ品目も含めてすべて出力。
    要確認品目は最も名前が近いマスター行の直後に挿入する。"""

    def make_agg_row(a, unknown=True):
        return {
            'genre':    a['genre'] if not unknown else (a['genre'] if a['genre'] != '要確認' else '要確認'),
            'baseName': a['baseName'],
            'g':        a['g'],
            'note':     '',
            'sp':       int(a['sp'])       if a['sp']       else '',
            'yokoSP':   int(a['yokoSP'])   if a['yokoSP']   else '',
            'mp':       int(a['mp'])       if a['mp']       else '',
            'mini':     int(a['mini'])     if a['mini']     else '',
            'takeuchi': int(a['takeuchi']) if a['takeuchi'] else '',
            'lotus':    int(a['lotus'])    if a['lotus']    else '',
            'unknown':  unknown,
        }

    # ① マスター行を順番通りに構築
    lookup = {(a['genre'] + '|' + a['baseName'] + '|' + a['g']): a for a in aggregated}
    master_rows = []
    used = set()

    for m in MASTER_DATA:
        name_norm = re.sub(r'\s+', ' ', m['name'].strip())
        key = m['genre'] + '|' + name_norm + '|' + m['g']
        used.add(key)
        a = lookup.get(key)
        master_rows.append({
            'genre':    m['genre'],
            'baseName': m['name'],
            'g':        m['g'],
            'note':     NOTE_MAP.get(m['name'], ''),
            'sp':       int(a['sp'])       if a and a['sp']       else '',
            'yokoSP':   int(a['yokoSP'])   if a and a['yokoSP']   else '',
            'mp':       int(a['mp'])       if a and a['mp']       else '',
            'mini':     int(a['mini'])     if a and a['mini']     else '',
            'takeuchi': int(a['takeuchi']) if a and a['takeuchi'] else '',
            'lotus':    int(a['lotus'])    if a and a['lotus']    else '',
            'unknown':  False,
        })

    # ② 要確認品目を収集
    unknowns = [a for a in aggregated if (a['genre'] + '|' + a['baseName'] + '|' + a['g']) not in used]

    # ③ 各要確認品目の挿入位置を決定（最も名前が近いマスター行の直後）
    def best_insert_idx(unknown_name, unknown_genre):
        """ジャンル一致を優先しつつ、名前が最も近いマスター行の直後を返す。"""
        u_ns = re.sub(r'\s+', '', unknown_name)

        def name_similar(mr_name):
            mr_ns = re.sub(r'\s+', '', mr_name)
            return (mr_ns == u_ns
                    or mr_ns in u_ns
                    or u_ns in mr_ns
                    or (len(u_ns) >= 4 and mr_ns[:4] == u_ns[:4]))

        # ① ジャンル一致 + 名前類似
        last_hit = -1
        for i, mr in enumerate(master_rows):
            if mr['genre'] == unknown_genre and name_similar(mr['baseName']):
                last_hit = i

        if last_hit >= 0:
            return last_hit

        # ② ジャンル不問で名前類似（ジャンル一致なし時のフォールバック）
        for i, mr in enumerate(master_rows):
            if name_similar(mr['baseName']):
                last_hit = i

        return last_hit  # -1 なら末尾

    # 挿入位置ごとにグループ化（同位置は元の順序を保つ）
    insert_map = {}   # {master_idx: [unknown agg items]}
    no_match   = []
    for a in unknowns:
        idx = best_insert_idx(a['baseName'], a['genre'])
        if idx >= 0:
            insert_map.setdefault(idx, []).append(a)
        else:
            no_match.append(a)

    # ④ マスター行 + 挿入行を合成
    final_rows = []
    for i, mr in enumerate(master_rows):
        final_rows.append(mr)
        for a in insert_map.get(i, []):
            final_rows.append(make_agg_row(a, unknown=True))

    # 名前で近い行が見つからなかった要確認は末尾へ
    for a in no_match:
        final_rows.append(make_agg_row(a, unknown=True))

    return final_rows


# ============================================================
# 売上合計計算
# ============================================================
def compute_total_sales(slips: list) -> int:
    """当日の全納品書の合計金額（税込）を返す"""
    total = 0
    for slip in slips:
        # トップレベルにある場合
        amount = slip.get('total_amount_including_tax') or slip.get('total_amount_with_tax')
        if amount is None:
            # body ネスト構造の場合
            body = slip.get('body', {}) or {}
            amount = body.get('total_amount_including_tax') or body.get('total_amount_with_tax') or 0
        try:
            total += int(float(amount or 0))
        except Exception:
            pass
    return total


def compute_shipping_total(slips: list) -> int:
    """当日の全納品書の送料合計を返す"""
    total = 0
    for slip in slips:
        items = slip.get('items') or slip.get('document_lines') or []
        for item in items:
            raw_name = item.get('name') or item.get('item_name') or ''
            if re.search(r'送料|手数料|配送|運賃', raw_name):
                amount = (item.get('total_amount_including_tax') or
                          item.get('subtotal_amount') or
                          item.get('total_amount') or 0)
                if not amount:
                    price = float(item.get('price') or item.get('unit_price') or 0)
                    qty   = float(item.get('quantity') or item.get('count') or 1)
                    amount = price * qty
                try:
                    total += int(float(amount or 0))
                except Exception:
                    pass
    return total


# ============================================================
# PDF 生成
# ============================================================
def generate_pdf(target_date_str: str, rows: list, output_path: str, total_sales: int = 0, shipping_total: int = 0, items: list = None):
    doc = SimpleDocTemplate(
        output_path, pagesize=B5,
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=10*mm, bottomMargin=10*mm,
    )

    title_style = ParagraphStyle(
        'title', fontName=FONT_NAME, fontSize=11, leading=16, spaceAfter=4*mm,
    )

    def to_int(v):
        try: return int(v)
        except: return 0

    total_sp   = sum(to_int(r['sp'])       for r in rows)
    total_yoko = sum(to_int(r['yokoSP'])   for r in rows)
    total_mp   = sum(to_int(r['mp'])       for r in rows)
    total_mini = sum(to_int(r['mini'])     for r in rows)
    total_take = sum(to_int(r['takeuchi']) for r in rows)
    total_lotu = sum(to_int(r['lotus'])    for r in rows)

    total_row_data = [
        '合計', '', '', '',
        str(total_sp)   if total_sp   else '',
        str(total_yoko) if total_yoko else '',
        str(total_mp)   if total_mp   else '',
        str(total_mini) if total_mini else '',
        str(total_take) if total_take else '',
        str(total_lotu) if total_lotu else '',
    ]

    VALUE_COLS = [('sp', 4), ('yokoSP', 5), ('mp', 6), ('mini', 7), ('takeuchi', 8), ('lotus', 9)]

    def _build_table(subset_rows, color_offset, include_total):
        td = [COL_HEADERS]
        for r in subset_rows:
            td.append([
                r['genre'], r['baseName'], r['g'], r['note'],
                r['sp'], r['yokoSP'], r['mp'], r['mini'],
                r['takeuchi'], r['lotus'],
            ])
        if include_total:
            td.append(total_row_data)

        n_rows = len(td)

        sc = [
            ('FONTNAME',      (0,0), (-1,-1), FONT_NAME),
            ('FONTSIZE',      (0,0), (-1,-1), 7),
            ('TOPPADDING',    (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING',   (0,0), (-1,-1), 2),
            ('RIGHTPADDING',  (0,0), (-1,-1), 2),
            ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
            # ヘッダー（左側）
            ('BACKGROUND', (0,0), (3,0),  HEADER_BG),
            ('TEXTCOLOR',  (0,0), (3,0),  HEADER_FG),
            # ヘッダー（パック列）
            ('BACKGROUND', (4,0), (-1,0), HEADER_PACK_BG),
            ('TEXTCOLOR',  (4,0), (-1,0), HEADER_PACK_FG),
            ('ALIGN',      (0,0), (-1,0), 'CENTER'),
            # 数値列は中央揃え
            ('ALIGN',      (4,1), (-1,-1), 'CENTER'),
            # 全行・全列の薄い罫線
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor('#C0C0C0')),
            # 外枠
            ('BOX',        (0,0), (-1,-1), 0.5, colors.grey),
            # ヘッダー下線
            ('LINEBELOW',  (0,0), (-1,0),  1,   colors.HexColor('#2C3E50')),
            # 備考 | SP 区切り縦線
            ('LINEAFTER',  (3,0), (3,-1),  1.5, colors.HexColor('#5D6D7E')),
        ]

        if include_total:
            sc.append(('BACKGROUND', (0, n_rows-1), (-1, n_rows-1), TOTAL_BG))
            sc.append(('FONTSIZE',   (0, n_rows-1), (-1, n_rows-1), 8))

        # 5行ごとの背景色 + 要確認オレンジ + 数値セル白背景
        for i, row in enumerate(subset_rows):
            row_idx = i + 1
            if row['unknown']:
                sc.append(('BACKGROUND', (0,row_idx), (-1,row_idx), UNKNOWN_COLOR))
            else:
                c = ROW_COLORS[((i + color_offset) // 5) % len(ROW_COLORS)]
                sc.append(('BACKGROUND', (0,row_idx), (-1,row_idx), c))
            for col_key, col_idx in VALUE_COLS:
                if row.get(col_key) not in ('', None, 0):
                    sc.append(('BACKGROUND', (col_idx,row_idx), (col_idx,row_idx), colors.white))

        # 5行ごとの区切り線
        for i in range(4, len(subset_rows), 5):
            sc.append(('LINEBELOW', (0, i+1), (-1, i+1), 0.8, colors.HexColor('#AAC4D0')))

        # タケウチ・ロテュス列：数字があるセルを赤枠＋赤下線で強調
        for i, row in enumerate(subset_rows):
            row_idx = i + 1
            try:
                has_take = int(row.get('takeuchi') or 0) > 0
            except (ValueError, TypeError):
                has_take = False
            try:
                has_lotu = int(row.get('lotus') or 0) > 0
            except (ValueError, TypeError):
                has_lotu = False
            if has_take:
                sc.append(('BOX', (8, row_idx), (8, row_idx), 1.5, colors.red))
            if has_lotu:
                sc.append(('BOX', (9, row_idx), (9, row_idx), 1.5, colors.red))
            if has_take or has_lotu:
                right_col = 9 if has_lotu else 8
                sc.append(('LINEBELOW', (0, row_idx), (right_col, row_idx), 1.0, colors.red))

        t = Table(td, colWidths=COL_WIDTHS, repeatRows=1)
        t.setStyle(TableStyle(sc))
        return t

    # エディブルフラワー先頭行を探してテーブルを分割（ページ先頭から始まるよう強制改ページ）
    split_idx = next(
        (i for i, r in enumerate(rows) if r['genre'] == 'エディブルフラワー'),
        None
    )

    if split_idx is None or split_idx == 0:
        table_items = [_build_table(rows, 0, include_total=True)]
    else:
        rows_part1 = rows[:split_idx]
        rows_part2 = rows[split_idx:]
        table_items = [
            _build_table(rows_part1, 0,         include_total=False),
            PageBreak(),
            _build_table(rows_part2, split_idx, include_total=True),
        ]

    def draw_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 9)
        canvas.drawCentredString(B5[0] / 2, 5 * mm, str(doc.page))
        canvas.restoreState()

    story = [Paragraph(f'農園 me!　パッキングリスト　{target_date_str}', title_style)] + table_items

    # ── サマリーページ ──────────────────────────────────────────
    SUMMARY_TARGETS = [
        {'label': 'マイクロリーフサラダミックス', 'genre': 'マイクロリーフ', 'name_key': 'サラダミックス'},
        {'label': 'チルドレンハーブミックス',     'genre': 'チルドレン',     'name_key': 'ハーブミックス'},
        {'label': 'チルドレン',                   'name_key': 'チルドレン',  'name_exclude': 'ハーブミックス'},
    ]

    def _to_int(v):
        try: return int(v)
        except: return 0

    summaries = []
    for tgt in SUMMARY_TARGETS:
        total_g = 0.0
        parts   = []
        for r in rows:
            if tgt['name_key'] not in r['baseName']:
                continue
            if tgt.get('name_exclude') and tgt['name_exclude'] in r['baseName']:
                continue
            if 'genre' in tgt and r['genre'] != tgt['genre']:
                continue
            g_val   = float(r['g']) if r['g'] else 0
            packs   = _to_int(r['sp']) + _to_int(r['yokoSP']) + _to_int(r['mp']) + _to_int(r['mini'])
            take_g  = _to_int(r['takeuchi'])
            lotu_g  = _to_int(r['lotus'])
            row_g   = packs * g_val + take_g + lotu_g
            if row_g == 0:
                continue
            total_g += row_g
            if packs > 0 and g_val > 0:
                g_disp = int(g_val) if g_val == int(g_val) else g_val
                parts.append(f'{g_disp}g×{packs}')
            if take_g > 0:
                parts.append(f'タケウチ{take_g}g')
            if lotu_g > 0:
                parts.append(f'ロテュス{lotu_g}g')
        if total_g > 0:
            t_disp = int(total_g) if total_g == int(total_g) else total_g
            summaries.append({
                'label':     tgt['label'],
                'total_g':   t_disp,
                'breakdown': '、'.join(parts),
            })

    # ── 最終ページは常に追加（挨拶・集計・備考欄）──────────────
    greet_style = ParagraphStyle(
        'greet', fontName=FONT_NAME, fontSize=11, leading=20, spaceAfter=3*mm,
    )
    item_style = ParagraphStyle(
        'sumitem', fontName=FONT_NAME, fontSize=12, leading=19, spaceAfter=3*mm,
    )
    section_style = ParagraphStyle(
        'section', fontName=FONT_NAME, fontSize=10, leading=16, spaceAfter=2*mm,
    )
    notice_style = ParagraphStyle(
        'notice', fontName=FONT_NAME, fontSize=11, leading=22, spaceAfter=2*mm,
    )

    greeting_text = (
        'まのくん、りさちゃん、こんにちは。<br/><br/>'
        'いつも丁寧なパック詰めをありがとう。<br/>'
        'お客様への感謝の気持ちを忘れずに、愛情をもってパック詰めをしてください。<br/>'
        'いつも応援しています。よろしくお願いします。'
    )

    story.append(PageBreak())
    story.append(Paragraph(greeting_text, greet_style))
    story.append(Spacer(1, 2*mm))

    # 集計行は数字がある品目のみ表示（・付き）
    for s in summaries:
        line = (
            f'・本日の{s["label"]}の出荷量は合計 {s["total_g"]}g になります。'
            f'（{s["breakdown"]}）'
        )
        story.append(Paragraph(line, item_style))

    # チルドレン＋チルドレンハーブミックス 合計行
    herb_total     = next((s['total_g'] for s in summaries if 'ハーブミックス' in s['label']), 0)
    children_total = next((s['total_g'] for s in summaries if s['label'] == 'チルドレン'), 0)
    if herb_total > 0 or children_total > 0:
        combined_total = herb_total + children_total
        combined_text = (
            f'・本日のチルドレンハーブミックスとチルドレンの出荷量の合計は '
            f'{combined_total}gとなります。'
            f'（チルドレン{children_total}g＋チルドレンハーブミックス{herb_total}g）'
        )
        story.append(Paragraph(combined_text, item_style))

    story.append(Spacer(1, 3*mm))

    # ── お客様別内訳テーブル（チルドレン・チルドレンハーブミックス）──
    if items:
        customer_detail = {}
        for it in items:
            genre = it.get('genre', '')
            name  = it.get('baseName', '')
            if not ((genre == 'チルドレン' and 'ハーブミックス' in name) or
                    (genre == 'その他' and 'チルドレン' in name and 'ハーブミックス' not in name)):
                continue
            cn      = it.get('customerName', '')
            g_val   = it.get('g', '')
            qty     = it.get('quantity', 1) * it.get('packCount', 1)
            g_num   = float(g_val) if g_val else 0
            total_g = int(g_num * qty) if g_num * qty == int(g_num * qty) else g_num * qty
            label   = f'{name} {g_val}g' if g_val else name
            key     = (cn, label)
            if key not in customer_detail:
                customer_detail[key] = {'packs': 0, 'total_g': 0}
            customer_detail[key]['packs']   += qty
            customer_detail[key]['total_g'] += total_g

        if customer_detail:
            story.append(Paragraph('【チルドレン・チルドレンハーブミックス　お客様別内訳】', notice_style))
            story.append(Spacer(1, 1*mm))
            c_rows = [['お客様名', '品目', '数量']]
            for (cn, label), v in customer_detail.items():
                c_rows.append([cn, label, f'{v["packs"]}パック（{v["total_g"]}g）'])
            c_table = Table(c_rows, colWidths=[55*mm, 65*mm, 36*mm])
            c_table.setStyle(TableStyle([
                ('FONTNAME',      (0, 0), (-1, -1), FONT_NAME),
                ('FONTSIZE',      (0, 0), (-1, 0),  9),
                ('FONTSIZE',      (0, 1), (-1, -1), 10),
                ('TOPPADDING',    (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING',   (0, 0), (-1, -1), 4),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
                ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
                ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor('#2C3E50')),
                ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.white),
                ('ALIGN',         (0, 0), (-1, 0),  'CENTER'),
                ('ALIGN',         (2, 1), (2, -1),  'CENTER'),
                ('INNERGRID',     (0, 0), (-1, -1), 0.3, colors.HexColor('#C0C0C0')),
                ('BOX',           (0, 0), (-1, -1), 0.5, colors.grey),
                ('LINEBELOW',     (0, 0), (-1, 0),  1,   colors.HexColor('#2C3E50')),
            ]))
            story.append(c_table)
            story.append(Spacer(1, 3*mm))

    # ── 特定品目の有無・売上確認 ──────────────────────────────────
    def _sum_packs(genre, name_contains):
        t = 0
        for r in rows:
            if r['genre'] == genre and name_contains in r['baseName']:
                for col in ['sp', 'yokoSP', 'mp', 'mini', 'takeuchi', 'lotus']:
                    try:
                        t += int(r[col])
                    except Exception:
                        pass
        return t

    salada_n   = _sum_packs('マイクロリーフ', 'サラダミックス')
    children_n = _sum_packs('その他', 'チルドレン')
    herb_n     = _sum_packs('チルドレン', 'ハーブミックス')

    salada_disp   = f'{salada_n}パック'   if salada_n   > 0 else 'ゼロです'
    children_disp = f'{children_n}パック' if children_n > 0 else 'ゼロです'
    herb_disp     = f'{herb_n}パック'     if herb_n     > 0 else 'ゼロです'

    notice_text = (
        f'サラダミックス：{salada_disp}　'
        f'チルドレン：{children_disp}　'
        f'チルドレンハーブミックス：{herb_disp}'
    )
    story.append(Paragraph(notice_text, notice_style))

    if total_sales > 0:
        story.append(Paragraph(f'本日の売上は {total_sales:,}円です', notice_style))
        if shipping_total > 0:
            product_total = total_sales - shipping_total
            story.append(Paragraph(
                f'うち送料は {shipping_total:,}円　商品のみの代金は {product_total:,}円です',
                notice_style,
            ))

    story.append(Spacer(1, 2*mm))

    # 備考欄
    story.append(Paragraph('備考・メモ欄', section_style))
    memo_table = Table(
        [[''], [''], [''], ['']],
        colWidths=[156*mm],
        rowHeights=[8*mm] * 4,
    )
    memo_table.setStyle(TableStyle([
        ('BOX',        (0, 0), (-1, -1), 1,   colors.black),
        ('INNERGRID',  (0, 0), (-1, -1), 0.3, colors.HexColor('#C0C0C0')),
        ('FONTNAME',   (0, 0), (-1, -1), FONT_NAME),
    ]))
    story.append(memo_table)

    # ── エディブルフラワーとその他のリーフ収穫量ページ（最終ページ）──────────────
    # ジャンル「エディブルフラワー」「その他」が対象（要確認行も含む）、合計パック数>0の品目を表示
    # 収穫数 = g値 × 合計パック数（sp+横SP+MP+ミニ+タケウチ+ロテュス）
    harvest_data = []
    for r in rows:
        if r['genre'] not in ('エディブルフラワー', 'その他'):
            continue
        total_packs = (to_int(r['sp']) + to_int(r['yokoSP']) +
                       to_int(r['mp']) + to_int(r['mini']) +
                       to_int(r['takeuchi']) + to_int(r['lotus']))
        if total_packs == 0:
            continue
        rinsu = float(r['g']) if r['g'] else 0
        harvest_count = round(rinsu * total_packs) if rinsu > 0 else total_packs
        rinsu_int = int(rinsu) if rinsu == int(rinsu) else rinsu
        formula = f'{rinsu_int}輪 × {total_packs}パック' if rinsu > 0 else f'{total_packs}パック'
        harvest_data.append([f'{r["genre"]} {r["baseName"]}', formula, str(harvest_count)])

    if harvest_data:
        story.append(PageBreak())
        harvest_title_style = ParagraphStyle(
            'harvest_title', fontName=FONT_NAME, fontSize=11, leading=16, spaceAfter=4*mm,
        )
        story.append(Paragraph('エディブルフラワーとその他のリーフ収穫量', harvest_title_style))

        h_table_data = [['品番・品名', '計算式', '収穫数']] + harvest_data
        harvest_table = Table(
            h_table_data,
            colWidths=[80*mm, 45*mm, 25*mm],
            repeatRows=1,
        )
        harvest_table.setStyle(TableStyle([
            ('FONTNAME',      (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE',      (0, 0), (-1, -1), 8),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor('#2C3E50')),
            ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.white),
            ('FONTSIZE',      (0, 0), (-1, 0),  9),
            ('ALIGN',         (0, 0), (-1, 0),  'CENTER'),
            ('ALIGN',         (1, 1), (1, -1),  'CENTER'),
            ('ALIGN',         (2, 1), (2, -1),  'CENTER'),
            ('INNERGRID',     (0, 0), (-1, -1), 0.3, colors.HexColor('#C0C0C0')),
            ('BOX',           (0, 0), (-1, -1), 0.5, colors.grey),
            ('LINEBELOW',     (0, 0), (-1, 0),  1,   colors.HexColor('#2C3E50')),
        ]))
        story.append(harvest_table)

    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    print(f'✅ PDF 生成完了: {output_path}')


# ============================================================
# メイン
# ============================================================
def main():
    if len(sys.argv) >= 2:
        target_date = sys.argv[1]
        try:
            datetime.strptime(target_date, '%Y-%m-%d')
        except ValueError:
            print('❌ 日付形式エラー。例: python3 misoca_packing_main.py 2026-03-29')
            sys.exit(1)
    else:
        target_date = date.today().strftime('%Y-%m-%d')

    date_str_jp = datetime.strptime(target_date, '%Y-%m-%d').strftime('%Y年%m月%d日')
    print('=' * 55)
    print(f'農園 me! パッキングリスト生成')
    print(f'対象日: {target_date}')
    print('=' * 55)

    slips      = fetch_delivery_slips(target_date, target_date)
    items      = process_slips(slips)
    aggregated = aggregate_items(items)
    rows       = build_rows_in_master_order(aggregated)

    unknown_count = sum(1 for r in rows if r['unknown'])
    print(f'集計完了: {len(rows)} 品目（うち要確認: {unknown_count} 件）')

    if unknown_count > 0:
        print('\n【要確認品目一覧】')
        for r in rows:
            if r['unknown']:
                vals = []
                if r['sp']:       vals.append(f'SP:{r["sp"]}')
                if r['yokoSP']:   vals.append(f'横SP:{r["yokoSP"]}')
                if r['mp']:       vals.append(f'MP:{r["mp"]}')
                if r['mini']:     vals.append(f'ミニ:{r["mini"]}')
                if r['takeuchi']: vals.append(f'タケウチ:{r["takeuchi"]}')
                if r['lotus']:    vals.append(f'ロテュス:{r["lotus"]}')
                print(f'  ジャンル:{r["genre"]} / 品名:{r["baseName"]} / g:{r["g"]} / {", ".join(vals) or "数量なし"}')

    total_sales    = compute_total_sales(slips)
    shipping_total = compute_shipping_total(slips)
    print(f'本日の売上合計（税込）: {total_sales:,}円')

    os.makedirs(ICLOUD_DIR, exist_ok=True)
    filename    = f'パッキングリスト_{target_date.replace("-","")}.pdf'
    output_path = os.path.join(ICLOUD_DIR, filename)

    generate_pdf(date_str_jp, rows, output_path, total_sales=total_sales, shipping_total=shipping_total, items=items)
    print(f'📂 保存先: {output_path}')

    if unknown_count > 0:
        print(f'\n⚠️  要確認品目が {unknown_count} 件あります（オレンジ行）。内容を確認してください。')


if __name__ == '__main__':
    main()
