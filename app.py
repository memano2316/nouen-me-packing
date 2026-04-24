#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
農園 me! パッキングリスト Web アプリ（Railway デプロイ用）
iPhone Safari からアクセス → 日付を入力 → PDF をブラウザで開く

環境変数:
  MISOCA_ACCESS_TOKEN   : Misoca アクセストークン
  MISOCA_REFRESH_TOKEN  : Misoca リフレッシュトークン
"""

import io
import json
import os
import re
from datetime import date, datetime

import requests
from flask import Flask, request, send_file, render_template_string
from reportlab.lib import colors
from reportlab.lib.pagesizes import B5
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

app = Flask(__name__)

# Misoca API との通信に使うセッション（TCP/TLS 接続を再利用して高速化）
misoca_session = requests.Session()

# ============================================================
# 設定
# ============================================================
TOKEN_FILE    = os.path.expanduser('~/.misoca_token.json')
API_BASE      = 'https://app.misoca.jp/api/v3'
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
# 商品マスター
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

COL_HEADERS   = ['ジャンル', '品名', 'g数', '備考', 'SP', '横SP', 'MP', 'ミニ', 'タケウチ', 'ロテュス']
COL_WIDTHS_MM = [24, 38, 12, 8, 11, 13, 11, 13, 13, 13]
COL_WIDTHS    = [w * mm for w in COL_WIDTHS_MM]


# ============================================================
# トークン管理（環境変数 → ファイル の順で取得）
# ============================================================

# dyno 起動中のトークンをメモリにキャッシュ（再起動までの間、毎回試し打ちを省略）
_token_cache: dict = {'access_token': None, 'refresh_token': None}


def get_valid_token() -> str:
    global _token_cache

    # ① キャッシュのトークンをまず試す
    if _token_cache['access_token']:
        r = misoca_session.get(
            f'{API_BASE}/delivery_slips?per_page=1&page=1',
            headers={'Authorization': f'Bearer {_token_cache["access_token"]}'},
            timeout=10,
        )
        if r.status_code == 200:
            return _token_cache['access_token']
        # 期限切れならキャッシュをクリアしてリフレッシュへ
        _token_cache['access_token'] = None

    # ② env var またはファイルから取得
    access_token = os.environ.get('MISOCA_ACCESS_TOKEN', '')
    refresh_tok  = _token_cache['refresh_token'] or os.environ.get('MISOCA_REFRESH_TOKEN', '')

    # ローカル開発用：環境変数がなければファイルから読む
    if not access_token and os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            data = json.load(f)
        access_token = data.get('access_token', '')
        if not refresh_tok:
            refresh_tok = data.get('refresh_token', '')

    if not access_token and not refresh_tok:
        raise Exception('MISOCA_ACCESS_TOKEN が設定されていません。Railway の環境変数を確認してください。')

    # ③ env var のトークンを試す（期限切れでなければそのまま使う）
    if access_token:
        r = misoca_session.get(
            f'{API_BASE}/delivery_slips?per_page=1&page=1',
            headers={'Authorization': f'Bearer {access_token}'},
            timeout=10,
        )
        if r.status_code == 200:
            _token_cache['access_token'] = access_token
            if refresh_tok:
                _token_cache['refresh_token'] = refresh_tok
            return access_token

    # ④ 401 → リフレッシュ試行
    if refresh_tok:
        resp = misoca_session.post(TOKEN_URL, data={
            'grant_type':    'refresh_token',
            'refresh_token': refresh_tok,
            'client_id':     CLIENT_ID,
            'client_secret': CLIENT_SECRET,
        }, timeout=10)
        if resp.status_code == 200:
            new_data = resp.json()
            # 新トークンをキャッシュに保持（dyno 再起動まで有効）
            _token_cache['access_token']  = new_data.get('access_token', '')
            _token_cache['refresh_token'] = new_data.get('refresh_token', refresh_tok)
            # ローカルならファイルも更新
            if os.path.exists(TOKEN_FILE):
                with open(TOKEN_FILE, 'w') as f:
                    json.dump(new_data, f, indent=2)
            return _token_cache['access_token']

    raise Exception('Misoca 認証エラー。トークンを更新してください。')


# ============================================================
# Misoca API：納品書取得
# ============================================================
def fetch_delivery_slips(start_date: str, end_date: str) -> list:
    token   = get_valid_token()
    headers = {'Authorization': f'Bearer {token}'}
    all_slips = []
    page, per_page = 1, 20

    while page <= 50:
        url = (f'{API_BASE}/delivery_slips'
               f'?issue_date_from={start_date}&issue_date_to={end_date}'
               f'&per_page={per_page}&page={page}')
        r = misoca_session.get(url, headers=headers, timeout=10)
        if r.status_code != 200:
            raise Exception(f'Misoca API エラー: HTTP {r.status_code}')
        data = r.json()
        if not isinstance(data, list) or len(data) == 0:
            break
        all_slips.extend(data)
        if len(data) < per_page:
            break
        page += 1

    # クライアント側で発行日を再フィルター（APIフィルターが不完全なため）
    return [s for s in all_slips if s.get('issue_date') == start_date]


# ============================================================
# データ処理
# ============================================================
def is_tokushu(customer_name: str) -> bool:
    return any(t in (customer_name or '') for t in TOKUSHU_CUSTOMERS)


KNOWN_GENRES = ['マイクロリーフ', 'エディブルフラワー', 'チルドレン', 'その他']


def parse_item_name(raw_name: str) -> dict:
    name = re.sub(r'\s+', ' ', (raw_name or '').strip())

    pack_count = 1
    m = re.search(r'[×x✕]\s*(\d+)', name, re.IGNORECASE)
    if m:
        pack_count = int(m.group(1)) or 1
        name = (name[:m.start()] + name[m.end():]).strip()

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

    qty_match = re.search(
        r'(\d+(?:\.\d+)?)\s*(?:g|輪入り?|本入り?|枚入り?|個入り?|枚|本|輪|個)$',
        name, re.IGNORECASE
    )
    g = ''
    if qty_match:
        g    = qty_match.group(1)
        name = name[:qty_match.start()].strip()

    return {'baseName': name, 'g': g, 'packCount': pack_count, 'detectedGenre': detected_genre}


def find_pack_from_master(base_name: str, g_val: str = '', detected_genre: str = '') -> dict:
    def hit(m):
        return {'genre': m['genre'], 'pack': m['pack'],
                'master_name': m['name'], 'unknown': False}

    def norm(s):
        return re.sub(r'\s+', ' ', s.strip()).replace('（', '(').replace('）', ')')

    def nsp(s):
        return re.sub(r'\s+', '', norm(s))

    normalized = norm(base_name) if base_name else ''
    nospace    = nsp(base_name)  if base_name else ''

    for m in MASTER_DATA:
        if norm(m['name']) == normalized and m['g'] == g_val:
            return hit(m)

    for m in MASTER_DATA:
        if nsp(m['name']) == nospace and m['g'] == g_val:
            return hit(m)

    if detected_genre and base_name:
        combined = nsp(detected_genre + base_name)
        for m in MASTER_DATA:
            if nsp(m['name']) == combined and m['g'] == g_val:
                return hit(m)

    if detected_genre and not base_name:
        genre_norm = norm(detected_genre)
        for m in MASTER_DATA:
            if norm(m['name']).startswith(genre_norm) and m['g'] == g_val:
                return hit(m)

    return {'genre': '要確認', 'pack': 'SP', 'master_name': '', 'unknown': True}


def process_slips(slips: list) -> list:
    results = []
    for slip in slips:
        customer = slip.get('recipient_name') or slip.get('contact_name') or ''
        tokushu  = is_tokushu(customer)
        items    = slip.get('items') or slip.get('document_lines') or []
        for item in items:
            raw_name = item.get('name') or item.get('item_name') or ''
            if re.search(r'送料|手数料|配送|運賃', raw_name):
                continue
            quantity = float(item.get('quantity') or item.get('count') or 1) or 1
            parsed   = parse_item_name(raw_name)
            master   = find_pack_from_master(parsed['baseName'], parsed['g'], parsed['detectedGenre'])

            if master['unknown']:
                use_name  = parsed['baseName']
                use_g     = parsed['g']
                use_genre = parsed['detectedGenre'] or '要確認'
                use_pack  = 'SP'
            else:
                use_name  = master['master_name']
                use_g     = parsed['g']
                use_genre = master['genre']
                use_pack  = master['pack']

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
            if is_takeuchi:
                agg[key]['takeuchi'] += g_num * qty
            else:
                agg[key]['lotus'] += g_num * qty
        else:
            p = item['pack']
            if p == 'SP':           agg[key]['sp']     += qty
            elif p == '横SP':       agg[key]['yokoSP'] += qty
            elif p == 'MP':         agg[key]['mp']     += qty
            elif p == 'ミニパック': agg[key]['mini']   += qty
            else:                   agg[key]['sp']     += qty
    return list(agg.values())


def build_rows_in_master_order(aggregated: list) -> list:
    def make_agg_row(a, unknown=True):
        return {
            'genre':    a['genre'],
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

    unknowns = [a for a in aggregated if (a['genre'] + '|' + a['baseName'] + '|' + a['g']) not in used]

    def best_insert_idx(unknown_name, unknown_genre):
        u_ns = re.sub(r'\s+', '', unknown_name)

        def name_similar(mr_name):
            mr_ns = re.sub(r'\s+', '', mr_name)
            return (mr_ns == u_ns
                    or mr_ns in u_ns
                    or u_ns in mr_ns
                    or (len(u_ns) >= 4 and mr_ns[:4] == u_ns[:4]))

        last_hit = -1
        for i, mr in enumerate(master_rows):
            if mr['genre'] == unknown_genre and name_similar(mr['baseName']):
                last_hit = i
        if last_hit >= 0:
            return last_hit

        for i, mr in enumerate(master_rows):
            if name_similar(mr['baseName']):
                last_hit = i
        return last_hit

    insert_map = {}
    no_match   = []
    for a in unknowns:
        idx = best_insert_idx(a['baseName'], a['genre'])
        if idx >= 0:
            insert_map.setdefault(idx, []).append(a)
        else:
            no_match.append(a)

    final_rows = []
    for i, mr in enumerate(master_rows):
        final_rows.append(mr)
        for a in insert_map.get(i, []):
            final_rows.append(make_agg_row(a, unknown=True))
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
# PDF 生成（BytesIO バッファ対応版）
# ============================================================
def generate_pdf(target_date_str: str, rows: list, output, total_sales: int = 0, shipping_total: int = 0, items: list = None) -> None:
    """output は ファイルパス文字列 または BytesIO バッファ"""
    doc = SimpleDocTemplate(
        output, pagesize=B5,
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
            ('BACKGROUND',    (0,0), (3,0),   HEADER_BG),
            ('TEXTCOLOR',     (0,0), (3,0),   HEADER_FG),
            ('BACKGROUND',    (4,0), (-1,0),  HEADER_PACK_BG),
            ('TEXTCOLOR',     (4,0), (-1,0),  HEADER_PACK_FG),
            ('ALIGN',         (0,0), (-1,0),  'CENTER'),
            ('ALIGN',         (4,1), (-1,-1), 'CENTER'),
            ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#C0C0C0')),
            ('BOX',           (0,0), (-1,-1), 0.5, colors.grey),
            ('LINEBELOW',     (0,0), (-1,0),  1,   colors.HexColor('#2C3E50')),
            ('LINEAFTER',     (3,0), (3,-1),  1.5, colors.HexColor('#5D6D7E')),
        ]

        if include_total:
            sc.append(('BACKGROUND', (0, n_rows-1), (-1, n_rows-1), TOTAL_BG))
            sc.append(('FONTSIZE',   (0, n_rows-1), (-1, n_rows-1), 8))

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

    # ── エディブルフラワー収穫数ページ（最終ページ）──────────────
    def to_int(v):
        try: return int(v)
        except: return 0

    harvest_data = []
    for r in rows:
        if r['genre'] != 'エディブルフラワー' or r.get('unknown', False):
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
        harvest_data.append([f'エディブルフラワー {r["baseName"]}', formula, str(harvest_count)])

    if harvest_data:
        story.append(PageBreak())
        harvest_title_style = ParagraphStyle(
            'harvest_title', fontName=FONT_NAME, fontSize=11, leading=16, spaceAfter=4*mm,
        )
        story.append(Paragraph('エディブルフラワー収穫数', harvest_title_style))
        h_table_data = [['品番・品名', '計算式', '収穫数']] + harvest_data
        harvest_table = Table(h_table_data, colWidths=[80*mm, 45*mm, 25*mm], repeatRows=1)
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


# ============================================================
# HTML テンプレート
# ============================================================
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
  <title>農園 me! パッキングリスト</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, 'Hiragino Sans', sans-serif;
      background: #f0f4f0;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      padding: 20px;
    }
    .card {
      background: white;
      border-radius: 16px;
      padding: 36px 24px;
      max-width: 400px;
      width: 100%;
      box-shadow: 0 4px 24px rgba(0,0,0,0.10);
      text-align: center;
    }
    .icon { font-size: 40px; margin-bottom: 8px; }
    h1 { font-size: 22px; color: #1E6B3C; margin-bottom: 4px; }
    .subtitle { color: #999; font-size: 13px; margin-bottom: 28px; }
    label { display: block; font-size: 14px; color: #555; margin-bottom: 8px; text-align: left; }
    input[type="date"] {
      width: 100%;
      padding: 14px 12px;
      font-size: 18px;
      border: 2px solid #ddd;
      border-radius: 10px;
      margin-bottom: 20px;
      color: #333;
      outline: none;
      transition: border-color 0.2s;
    }
    input[type="date"]:focus { border-color: #1E6B3C; }
    button {
      width: 100%;
      padding: 16px;
      font-size: 18px;
      font-weight: bold;
      background: #1E6B3C;
      color: white;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      transition: background 0.2s, transform 0.1s;
    }
    button:active { background: #155a30; transform: scale(0.98); }
    button:disabled { background: #aaa; cursor: not-allowed; }
    .status { margin-top: 16px; font-size: 14px; color: #888; min-height: 20px; }
    .error { color: #e74c3c !important; }
  </style>
</head>
<body>
  <div class="card">
    <div class="icon">🌿</div>
    <h1>農園 me!</h1>
    <p class="subtitle">パッキングリスト生成</p>
    <form id="form" action="/generate" method="post" target="_blank">
      <label>対象日</label>
      <input type="date" name="date" id="date" value="{{ today }}" required>
      <button type="submit" id="btn">PDF を生成する</button>
    </form>
    <p class="status" id="status"></p>
  </div>
  <script>
    document.getElementById('form').addEventListener('submit', function() {
      var btn = document.getElementById('btn');
      btn.disabled = true;
      btn.textContent = '生成中...';
      document.getElementById('status').textContent = 'Misoca からデータを取得しています…';
      setTimeout(function() {
        btn.disabled = false;
        btn.textContent = 'PDF を生成する';
        document.getElementById('status').textContent = '';
      }, 15000);
    });
  </script>
</body>
</html>"""


# ============================================================
# Flask ルート
# ============================================================
@app.route('/')
def index():
    today = date.today().strftime('%Y-%m-%d')
    return render_template_string(HTML_TEMPLATE, today=today)


@app.route('/generate', methods=['POST'])
def generate():
    target_date = request.form.get('date', date.today().strftime('%Y-%m-%d'))
    try:
        datetime.strptime(target_date, '%Y-%m-%d')
    except ValueError:
        return '日付の形式が正しくありません', 400

    try:
        slips           = fetch_delivery_slips(target_date, target_date)
        items           = process_slips(slips)
        aggregated      = aggregate_items(items)
        rows            = build_rows_in_master_order(aggregated)
        total_sales     = compute_total_sales(slips)
        shipping_total  = compute_shipping_total(slips)
        date_str_jp     = datetime.strptime(target_date, '%Y-%m-%d').strftime('%Y年%m月%d日')

        buf = io.BytesIO()
        generate_pdf(date_str_jp, rows, buf, total_sales=total_sales, shipping_total=shipping_total, items=items)
        buf.seek(0)

        filename = f'パッキングリスト_{target_date.replace("-", "")}.pdf'
        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=False,
            download_name=filename,
        )
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        return f'<pre style="color:red;padding:20px;">エラー: {e}\n\n{tb}</pre>', 500
    except BaseException as e:
        import traceback
        tb = traceback.format_exc()
        return f'<pre style="color:red;padding:20px;">致命的エラー: {e}\n\n{tb}</pre>', 500


@app.route('/debug')
def debug():
    import traceback
    today = date.today().strftime('%Y-%m-%d')
    try:
        token = get_valid_token()
        slips = fetch_delivery_slips(today, today)
        return f'<pre>token OK\nslips: {len(slips)} 件</pre>'
    except Exception as e:
        return f'<pre style="color:red;">{e}\n{traceback.format_exc()}</pre>', 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
