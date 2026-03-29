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
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

app = Flask(__name__)

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
    {'genre': 'エディブルフラワー', 'name': 'カレンデュラ イエロー', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'カレンデュラ  オレンジ', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'カレンデュラ ミックス', 'g': '5', 'pack': '横SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー ブルー', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー ピンク', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー レッド', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー バイオレット', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー チョコレート', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': 'コーンフラワー ミックス', 'g': '15', 'pack': 'SP'},
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
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド  イエロー', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド オレンジ', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド サンライズ', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド レッドフォックス', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド ストロベリーアンティーク', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド ミックス', 'g': '15', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'マリーゴールド ミックス(ミニ)', 'g': '4', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'ミニマリーゴールド', 'g': '30', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '黄花コスモス', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': '黄花コスモス(ミニ)', 'g': '4', 'pack': 'ミニパック'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタス ホワイト', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタス レッド', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタス ピンク', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタス  バイオレット', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ペンタス ミックス', 'g': '8', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'ベゴニアミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'その他', 'name': 'ベゴニアのつぼみ', 'g': '25', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニア レッド', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニア　ホワイト', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニア　ピンク', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きベゴニアミックス', 'g': '15', 'pack': 'SP'},
    {'genre': 'エディブルフラワー', 'name': '八重咲きペチュニア', 'g': '12', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナ レッド', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナ ホワイト', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナピンク', 'g': '10', 'pack': 'MP'},
    {'genre': 'エディブルフラワー', 'name': 'バーベナ バイオレット', 'g': '10', 'pack': 'MP'},
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
def get_valid_token() -> str:
    access_token = os.environ.get('MISOCA_ACCESS_TOKEN', '')
    refresh_tok  = os.environ.get('MISOCA_REFRESH_TOKEN', '')

    # ローカル開発用：環境変数がなければファイルから読む
    if not access_token and os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            data = json.load(f)
        access_token = data.get('access_token', '')
        refresh_tok  = data.get('refresh_token', '')

    if not access_token:
        raise Exception('MISOCA_ACCESS_TOKEN が設定されていません。Railway の環境変数を確認してください。')

    # 試し打ち
    r = requests.get(
        f'{API_BASE}/delivery_slips?per_page=1&page=1',
        headers={'Authorization': f'Bearer {access_token}'},
    )
    if r.status_code == 200:
        return access_token

    # 401 → リフレッシュ試行
    if r.status_code == 401 and refresh_tok:
        resp = requests.post(TOKEN_URL, data={
            'grant_type':    'refresh_token',
            'refresh_token': refresh_tok,
            'client_id':     CLIENT_ID,
            'client_secret': CLIENT_SECRET,
        })
        if resp.status_code == 200:
            new_data = resp.json()
            # ローカルならファイルも更新
            if os.path.exists(TOKEN_FILE):
                with open(TOKEN_FILE, 'w') as f:
                    json.dump(new_data, f, indent=2)
            return new_data.get('access_token', '')

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
        r = requests.get(url, headers=headers)
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
# PDF 生成（BytesIO バッファ対応版）
# ============================================================
def generate_pdf(target_date_str: str, rows: list, output) -> None:
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

    table_data = [COL_HEADERS]
    for r in rows:
        table_data.append([
            r['genre'], r['baseName'], r['g'], r['note'],
            r['sp'], r['yokoSP'], r['mp'], r['mini'],
            r['takeuchi'], r['lotus'],
        ])
    table_data.append([
        '合計', '', '', '',
        str(total_sp)   if total_sp   else '',
        str(total_yoko) if total_yoko else '',
        str(total_mp)   if total_mp   else '',
        str(total_mini) if total_mini else '',
        str(total_take) if total_take else '',
        str(total_lotu) if total_lotu else '',
    ])

    total_rows = len(table_data)

    style_cmds = [
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
        ('BACKGROUND',    (0, total_rows-1), (-1, total_rows-1), TOTAL_BG),
        ('FONTSIZE',      (0, total_rows-1), (-1, total_rows-1), 8),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#C0C0C0')),
        ('BOX',           (0,0), (-1,-1), 0.5, colors.grey),
        ('LINEBELOW',     (0,0), (-1,0),  1,   colors.HexColor('#2C3E50')),
        ('LINEAFTER',     (3,0), (3,-1),  1.5, colors.HexColor('#5D6D7E')),
    ]

    VALUE_COLS = [('sp', 4), ('yokoSP', 5), ('mp', 6), ('mini', 7), ('takeuchi', 8), ('lotus', 9)]
    for i, row in enumerate(rows):
        row_idx = i + 1
        if row['unknown']:
            style_cmds.append(('BACKGROUND', (0,row_idx), (-1,row_idx), UNKNOWN_COLOR))
        else:
            c = ROW_COLORS[(i // 5) % len(ROW_COLORS)]
            style_cmds.append(('BACKGROUND', (0,row_idx), (-1,row_idx), c))
        for col_key, col_idx in VALUE_COLS:
            if row.get(col_key) not in ('', None, 0):
                style_cmds.append(('BACKGROUND', (col_idx,row_idx), (col_idx,row_idx), colors.white))

    for i in range(4, len(rows), 5):
        style_cmds.append(('LINEBELOW', (0, i+1), (-1, i+1), 0.8, colors.HexColor('#AAC4D0')))

    table = Table(table_data, colWidths=COL_WIDTHS, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))

    story = [
        Paragraph(f'農園 me!　パッキングリスト　{target_date_str}', title_style),
        table,
    ]
    doc.build(story)


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
        slips      = fetch_delivery_slips(target_date, target_date)
        items      = process_slips(slips)
        aggregated = aggregate_items(items)
        rows       = build_rows_in_master_order(aggregated)
        date_str_jp = datetime.strptime(target_date, '%Y-%m-%d').strftime('%Y年%m月%d日')

        buf = io.BytesIO()
        generate_pdf(date_str_jp, rows, buf)
        buf.seek(0)

        filename = f'パッキングリスト_{target_date.replace("-", "")}.pdf'
        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=False,
            download_name=filename,
        )
    except Exception as e:
        return f'<p style="color:red;font-family:sans-serif;padding:20px;">エラー: {e}</p>', 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
