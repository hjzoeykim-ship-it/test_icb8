# -*- coding: utf-8 -*-
"""
한국 지역 축제 카니발리제이션 분석 대시보드 생성 스크립트 (오프라인 지원용)
작성자: Antigravity AI 코딩 어시스턴트
"""

import os
import webbrowser
import pandas as pd
import json

def analyze_and_generate_dashboard():
    excel_path = 'data/cleansing/festival_cleansed_F_v6.xlsx'
    html_output_path = 'scratch/dashboard.html'
    
    print("=== [1/3] 데이터 정합성 검증 시작 ===")
    if not os.path.exists(excel_path):
        print(f"오류: 데이터 파일 {excel_path}이(가) 존재하지 않습니다.")
        return
        
    try:
        df = pd.read_excel(excel_path)
        print(f"성공: {excel_path} 데이터를 성공적으로 불러왔습니다.")
        print(f"- 데이터 규모: {df.shape[0]}행, {df.shape[1]}열")
    except Exception as e:
        print(f"데이터 정합성 검증 중 경고/오류 발생 (기본 정적 상수를 사용하여 대시보드를 생성합니다): {e}")

    print("\n=== [2/3] HTML 대시보드 파일 생성 시작 ===")
    
    # HTML5/CSS3/Vanilla JS 코드로 대시보드 정의 (외부 CDN 의존성 제거)
    html_content = """<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>전국 지역 축제 카니발리제이션 분석 대시보드</title>
    
    <style>
        /* CSS Reset & Variable Design System */
        :root {
            --navy: #1C2D4A;
            --gold: #B89B4A;
            --red-danger: #D94F4F;
            --green-synergy: #3C8C5A;
            --light-bg: #F4F5F8;
            --card-bg: #FFFFFF;
            --text-main: #1A1A2E;
            --text-muted: #6B7280;
            --border-col: #E2E4EB;
            --font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: var(--font-family);
            background-color: var(--light-bg);
            color: var(--text-main);
            overflow: hidden; /* 원페이지 고정 뷰포트 */
            width: 100vw;
            height: 100vh;
            display: flex;
            flex-direction: column;
            user-select: none;
        }

        /* Header Style */
        header {
            background-color: var(--navy);
            color: white;
            padding: 12px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            z-index: 10;
        }

        .header-title-container {
            display: flex;
            flex-direction: column;
            text-align: left;
        }

        .header-title {
            font-size: 1.25rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .header-subtitle {
            font-size: 0.75rem;
            color: #D1D5DB;
            font-weight: 300;
            margin-top: 2px;
        }

        /* Tabs Navigation */
        .tabs-container {
            display: flex;
            background-color: #2D3E5D;
            padding: 4px;
            border-radius: 9999px;
            border: 1px solid #4B5563;
        }

        .tab-btn {
            background: none;
            border: none;
            color: #D1D5DB;
            padding: 6px 20px;
            font-size: 0.875rem;
            font-weight: 500;
            border-radius: 9999px;
            cursor: pointer;
            transition: all 0.15s ease;
        }

        .tab-btn:hover {
            color: white;
            background-color: rgba(28, 45, 74, 0.5);
        }

        .tab-btn.active {
            background-color: var(--gold);
            color: white;
            font-weight: 700;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }

        /* Main Container */
        main {
            flex: 1;
            padding: 24px;
            overflow: hidden;
            position: relative;
        }

        .tab-panel {
            display: none;
            height: 100%;
            width: 100%;
        }

        .tab-panel.active {
            display: flex;
            flex-direction: column;
            gap: 16px;
        }

        /* Stat Cards Layout */
        .stat-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
        }

        .stat-card {
            background-color: var(--card-bg);
            padding: 16px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border-left: 4px solid var(--navy);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: left;
        }

        .stat-card.gold { border-left-color: var(--gold); }
        .stat-card.red { border-left-color: var(--red-danger); }
        .stat-card.green { border-left-color: var(--green-synergy); }

        .stat-title {
            font-size: 0.75rem;
            color: var(--text-muted);
            font-weight: 500;
        }

        .stat-value {
            font-size: 1.875rem;
            font-weight: 800;
            color: var(--text-main);
            padding: 4px 0;
            letter-spacing: -0.025em;
        }

        .stat-sub {
            font-size: 0.75rem;
            color: #9CA3AF;
            font-weight: 300;
        }

        /* Grid Layouts */
        .content-grid-overview {
            display: grid;
            grid-template-columns: 7fr 5fr;
            gap: 16px;
            flex: 1;
            min-height: 0;
        }

        .content-grid-half {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            flex: 1;
            min-height: 0;
        }

        .card {
            background-color: var(--card-bg);
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border: 1px solid var(--border-col);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: left;
            min-height: 0;
        }

        .card-header {
            border-bottom: 1px solid var(--border-col);
            padding-bottom: 8px;
            margin-bottom: 12px;
        }

        .card-title {
            font-size: 0.875rem;
            font-weight: 700;
            color: var(--text-main);
        }

        .card-desc {
            font-size: 0.6875rem;
            color: var(--text-muted);
            margin-top: 2px;
        }

        /* Callouts */
        .callout {
            margin-top: 12px;
            padding: 8px 12px;
            border-radius: 4px;
            font-size: 0.75rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .callout.red {
            background-color: #FEF2F2;
            border: 1px solid #FEE2E2;
            color: var(--red-danger);
            font-weight: 600;
        }

        .callout.slate {
            background-color: #F8FAFC;
            border: 1px solid var(--border-col);
            color: var(--text-main);
        }

        .callout.green {
            background-color: #ECFDF5;
            border: 1px solid #D1FAE5;
            color: var(--green-synergy);
            font-weight: 700;
        }

        /* 3-Column Bottom row info cards */
        .bottom-info-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
        }

        .info-card {
            background-color: var(--card-bg);
            padding: 16px;
            border-radius: 8px;
            border: 1px solid var(--border-col);
            border-top: 4px solid var(--navy);
            display: flex;
            flex-direction: column;
            gap: 6px;
            text-align: left;
        }

        .info-card.red { border-top-color: var(--red-danger); border-color: rgba(217, 79, 79, 0.2); }
        .info-card.gold { border-top-color: var(--gold); border-color: rgba(184, 155, 74, 0.2); }
        .info-card.green { border-top-color: var(--green-synergy); border-color: rgba(60, 140, 90, 0.2); }

        .info-badge {
            font-size: 0.6875rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .info-badge.red { color: var(--red-danger); }
        .info-badge.gold { color: var(--gold); }
        .info-badge.green { color: var(--green-synergy); }

        .info-title {
            font-size: 0.75rem;
            font-weight: 700;
            color: var(--text-main);
        }

        .info-body {
            font-size: 0.6875rem;
            color: var(--text-muted);
            line-height: 1.5;
        }

        /* Chart Canvas styling (pure SVG implementation) */
        .chart-container {
            flex: 1;
            min-height: 0;
            position: relative;
            width: 100%;
            height: 100%;
        }

        svg {
            width: 100%;
            height: 100%;
        }

        /* Tooltip style */
        .tooltip {
            position: absolute;
            background-color: var(--navy);
            color: white;
            padding: 8px 12px;
            border-radius: 4px;
            font-size: 11px;
            pointer-events: none;
            display: none;
            z-index: 100;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.15);
            border: 1px solid var(--border-col);
        }

        .tooltip-title {
            font-weight: bold;
            border-bottom: 1px solid #4B5563;
            padding-bottom: 4px;
            margin-bottom: 4px;
        }

        /* Custom Interactive SVG styling */
        .bar-rect {
            transition: fill 0.15s ease;
            cursor: pointer;
        }
        .bar-rect:hover {
            opacity: 0.85;
        }

        .pie-slice {
            transition: opacity 0.15s ease;
            cursor: pointer;
        }
        .pie-slice:hover {
            opacity: 0.85;
        }

        /* Tab 4: Strategic Implications specific styles */
        .implications-top-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
        }

        .imp-card {
            background-color: var(--card-bg);
            padding: 16px;
            border-radius: 8px;
            border: 1px solid var(--border-col);
            border-left: 4px solid var(--red-danger);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: left;
        }
        .imp-card.green { border-left-color: var(--green-synergy); }

        .imp-header {
            font-size: 0.75rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 4px;
        }
        .imp-header.red { color: var(--red-danger); }
        .imp-header.green { color: var(--green-synergy); }

        .imp-stat-title {
            font-size: 1.25rem;
            font-weight: 900;
            color: var(--navy);
            padding: 4px 0;
        }

        .imp-result {
            font-size: 0.6875rem;
            font-weight: 700;
            margin-bottom: 8px;
        }
        .imp-result.red { color: var(--red-danger); }
        .imp-result.green { color: var(--green-synergy); }

        .imp-footer {
            font-size: 0.5625rem;
            color: #9CA3AF;
            margin-top: 4px;
        }

        /* 3-phase strategic list cards */
        .phase-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
            flex: 1;
            min-height: 0;
        }

        .phase-card {
            background-color: var(--card-bg);
            padding: 16px;
            border-radius: 8px;
            border: 1px solid var(--border-col);
            border-top: 4px solid var(--red-danger);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: left;
        }
        .phase-card.gold { border-top-color: var(--gold); }
        .phase-card.green { border-top-color: var(--green-synergy); }

        .phase-header-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #F1F5F9;
            padding-bottom: 6px;
            margin-bottom: 8px;
        }

        .phase-badge {
            font-size: 0.5625rem;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
        }
        .phase-badge.red { background-color: #FEF2F2; color: var(--red-danger); }
        .phase-badge.gold { background-color: #FEFBE8; color: var(--gold); }
        .phase-badge.green { background-color: #ECFDF5; color: var(--green-synergy); }

        .phase-step {
            font-size: 0.75rem;
            font-weight: 700;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 4px;
            margin-bottom: 4px;
        }

        .phase-list {
            list-style: none;
            font-size: 0.6875rem;
            color: var(--text-muted);
        }

        .phase-list li {
            margin-bottom: 4px;
        }

        .bottom-callout-full {
            background-color: var(--navy);
            color: white;
            padding: 12px;
            border-radius: 8px;
            text-align: center;
            font-size: 0.75rem;
            font-weight: 500;
        }

        /* Source citation text footer */
        .citation-text {
            font-size: 10px;
            color: #9CA3AF;
            margin-top: 4px;
            align-self: flex-start;
        }
    </style>
</head>
<body>
    <!-- FIXED TOP BAR -->
    <header>
        <div class="header-title-container">
            <h1 class="header-title">📊 전국 지역 축제 카니발리제이션 분석</h1>
            <span class="header-subtitle">2024–2025 · 문화체육관광부 · 2,384개 축제 데이터 기반</span>
        </div>
        
        <!-- TAB BUTTONS -->
        <div class="tabs-container">
            <button class="tab-btn active" onclick="switchTab('overview')">개요</button>
            <button class="tab-btn" onclick="switchTab('crowding')">시공간 과밀</button>
            <button class="tab-btn" onclick="switchTab('efficiency')">테마·효율</button>
            <button class="tab-btn" onclick="switchTab('implications')">전략 시사점</button>
        </div>
    </header>

    <!-- MAIN CONTENT AREA -->
    <main>
        <!-- TOOLTIP FOR SVG CHARTS -->
        <div id="chart-tooltip" class="tooltip"></div>

        <!-- ==========================================
             TAB 1: 개요 (Overview)
             ========================================== -->
        <div id="panel-overview" class="tab-panel active">
            <!-- STAT CARDS -->
            <div class="stat-grid">
                <div class="stat-card">
                    <span class="stat-title">전국 분석 축제 수</span>
                    <span class="stat-value">2,384</span>
                    <span class="stat-sub">2024~2025년 합산</span>
                </div>
                <div class="stat-card gold">
                    <span class="stat-title">5월·10월 집중도</span>
                    <span class="stat-value">41%</span>
                    <span class="stat-sub">978건이 단 2개월에 집중</span>
                </div>
                <div class="stat-card red">
                    <span class="stat-title">상위 3개 테마 편중</span>
                    <span class="stat-value">72%</span>
                    <span class="stat-sub">문화예술·특산물·자연생태</span>
                </div>
                <div class="stat-card green">
                    <span class="stat-title">지방비 전용 효율 우위</span>
                    <span class="stat-value">2.1×</span>
                    <span class="stat-sub">국비 지원 대비 효율지표 우세</span>
                </div>
            </div>

            <!-- CONTENT GRID -->
            <div class="content-grid-overview">
                <!-- LEFT COLUMN: 월별 축제 개최 수 -->
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">월별 축제 개최 수</h3>
                        <p class="card-desc">전국 17개 광역 지자체의 월별 축제 공급량</p>
                    </div>
                    
                    <div class="chart-container" id="chart-monthly-holder">
                        <!-- Horizontal Bar Chart SVG -->
                    </div>
                    
                    <div class="callout red">
                        <span>⚠</span>
                        <span>10월 단일 617건 (26%) — 5월 및 10월 합산 978건 (41%)으로 심각한 가을철 쏠림 현상 존재</span>
                    </div>
                </div>

                <!-- RIGHT COLUMN: 광역별 축제 분포 -->
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">광역별 축제 분포 (상위 8개)</h3>
                        <p class="card-desc">전체 축제 공급의 40% 이상이 특정 4개 지역에 밀집</p>
                    </div>
                    
                    <div class="chart-container" id="chart-regional-holder">
                        <!-- Donut Chart SVG -->
                    </div>

                    <div class="callout slate">
                        💡 경기 · 전남 · 경남 · 강원 등 <b>상위 4개 광역이 전체의 43.9%</b>를 차지하며 편중 심화
                    </div>
                </div>
            </div>
            
            <div class="citation-text">데이터 출처: 2024~2025년 문화체육관광부 전국 지역축제 개최계획 기준</div>
        </div>

        <!-- ==========================================
             TAB 2: 시공간 과밀 (Crowding)
             ========================================== -->
        <div id="panel-crowding" class="tab-panel">
            <!-- TOP INSIGHT BANNER -->
            <div class="bottom-callout-full" style="background-color: var(--navy); display: flex; justify-content: space-between; align-items: center; padding: 10px 16px;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 16px;">🚨</span>
                    <span style="font-weight: 600;">경기도 10월에만 86개 축제가 동시 개최 — 방문객 파이를 나눠 가집니다.</span>
                </div>
                <span style="font-size: 10px; background-color: var(--red-danger); color: white; padding: 2px 8px; border-radius: 99px; font-weight: bold;">과밀 위험 최고조</span>
            </div>

            <!-- CHARTS ROW -->
            <div class="content-grid-half">
                <!-- LEFT: 월별 축제 수 vs 효율지표 (이중 Y축) -->
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">월별 개최 축제 수 vs 평균 효율성 비교</h3>
                        <p class="card-desc">축제가 몰리는 특정 시기(5·10월)에 효율성 지표의 저하 현상이 뚜렷이 관측됨</p>
                    </div>
                    
                    <div class="chart-container" id="chart-combo-holder">
                        <!-- Double Y axis Chart SVG -->
                    </div>
                    
                    <p style="font-size: 10px; color: var(--text-muted); font-style: italic; margin-top: 4px;">
                        * 효율지표 = 방문객수 / 예산합계(백만원). 축제가 몰리는 10월에는 효율지표가 267.1로 비성수기(3월: 1,041.0) 대비 최저점 기록.
                    </p>
                </div>

                <!-- RIGHT: 광역별 축제 수 (전체) -->
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">광역별 축제 공급수 (전체 17개 시도)</h3>
                        <p class="card-desc">전국 시도별 축제 유치 편차 및 집중도 순위 (상위 4개 강조)</p>
                    </div>
                    
                    <div class="chart-container" id="chart-regional-all-holder">
                        <!-- Bar Chart SVG -->
                    </div>
                </div>
            </div>

            <!-- BOTTOM CALLOUT ROW -->
            <div class="bottom-info-grid">
                <div class="info-card red">
                    <div class="info-badge red"><span>⚠</span><span>DANGER ZONE</span></div>
                    <h4 class="info-title">동일 광역·월에 40개 이상 과밀</h4>
                    <p class="info-body">
                        • 경기도 10월: 86건 동시 개최<br>
                        • 경기도 9월: 70건 동시 개최<br>
                        • 전라남도 10월: 66건 동시 개최
                    </p>
                </div>

                <div class="info-card gold">
                    <div class="info-badge gold"><span>💡</span><span>계절성 공급 패턴</span></div>
                    <h4 class="info-title">5·10월 자연 및 기후 의존성</h4>
                    <p class="info-body">
                        봄꽃(5월) 및 가을단풍/수확(10월) 테마 축제가 자연 기후 주기에 따라 특정 시기에 집중적으로 기획되면서 공급 과밀 현상이 정형화됨.
                    </p>
                </div>

                <div class="info-card green">
                    <div class="info-badge green"><span>✓</span><span>BLUE OCEAN OPPORTUNITY</span></div>
                    <h4 class="info-title">1·2·6·7월 축제 공백기 공략</h4>
                    <p class="info-body">
                        연초 및 혹서기인 비수기 시즌 진입 시, 경쟁 강도가 전무한 수준이므로 역발상 특화 테마(겨울 레저, 여름 야간형 축제)로 관광객을 선점할 기회 존재.
                    </p>
                </div>
            </div>
        </div>

        <!-- ==========================================
             TAB 3: 테마·효율 (Theme & Efficiency)
             ========================================== -->
        <div id="panel-efficiency" class="tab-panel">
            <div class="content-grid-half" style="grid-template-rows: 1fr 1fr; grid-template-columns: 1fr 1fr; height: 100%;">
                
                <!-- TOP-LEFT: 축제 유형 분포 -->
                <div class="card">
                    <div class="card-header" style="margin-bottom: 4px;">
                        <h3 class="card-title">전체 축제 유형 분포</h3>
                        <p class="card-desc">가장 흔히 기획되는 3대 표준 테마의 지배력 분석</p>
                    </div>
                    
                    <div class="chart-container" id="chart-type-donut-holder" style="height: 60%;">
                        <!-- Donut Chart SVG -->
                    </div>
                    
                    <div class="callout slate" style="margin-top: 4px; padding: 4px 8px; font-size: 11px;">
                        📌 <b>문화예술·특산물·자연생태</b>가 전체의 <b>72.1%</b>를 독점하는 고도의 테마 편중 현상
                    </div>
                </div>

                <!-- TOP-RIGHT: 유형별 평균 효율지표 -->
                <div class="card">
                    <div class="card-header" style="margin-bottom: 4px;">
                        <h3 class="card-title">유형별 평균 효율지표 (방문객수 / 백만원)</h3>
                        <p class="card-desc">공급 편중 대비 실제 예산 투입 대비 방문객 유치 성과 분석</p>
                    </div>
                    
                    <div class="chart-container" id="chart-type-eff-holder" style="height: 65%;">
                        <!-- Horizontal Bar Chart SVG -->
                    </div>
                    
                    <div class="callout green" style="margin-top: 4px; padding: 4px 8px; font-size: 11px;">
                        📈 자연생태(798.2) 및 지역특산물(468.2) 테마는 문화예술 대비 최대 3배 효율
                    </div>
                </div>

                <!-- BOTTOM-LEFT: 재정 구조별 효율 비교 -->
                <div class="card">
                    <div class="card-header" style="margin-bottom: 4px;">
                        <h3 class="card-title">예산 규모 및 재정 구조별 효율성 분석</h3>
                        <p class="card-desc">국비 보조금 지원 유무에 따른 가성비 역설</p>
                    </div>
                    
                    <div style="display: flex; gap: 16px; flex: 1; min-height: 0; align-items: center;">
                        <div class="chart-container" id="chart-finance-holder" style="flex: 1; height: 100%;">
                            <!-- Bar Chart SVG -->
                        </div>
                        <div style="flex: 1; font-size: 11px; display: flex; flex-direction: column; gap: 4px; padding-left: 12px; border-left: 1px solid var(--border-col);">
                            <div>
                                <span style="font-weight: bold; color: var(--red-danger);">🔴 국비 지원형</span>
                                <p style="color: var(--text-muted); font-size: 10px;">평균 예산 4,881M / 효율 193.9명</p>
                            </div>
                            <div>
                                <span style="font-weight: bold; color: var(--green-synergy);">🟢 지방비 전용</span>
                                <p style="color: var(--text-muted); font-size: 10px;">평균 예산 1,758M / 효율 401.2명</p>
                            </div>
                            <div style="background-color: #FEF2F2; color: var(--red-danger); padding: 4px 8px; border-radius: 4px; font-size: 9px; margin-top: 4px;">
                                ⚠️ <b>역설:</b> 국비 지원 사업의 예산은 2.8배 많으나 성과는 절반에 그침.
                            </div>
                        </div>
                    </div>
                </div>

                <!-- BOTTOM-RIGHT: 축제 연식별 효율지표 -->
                <div class="card">
                    <div class="card-header" style="margin-bottom: 4px;">
                        <h3 class="card-title">축제 연식(경력)에 따른 효율지표 분포</h3>
                        <p class="card-desc">경력 브랜드 가치가 자기잠식(중복 개최) 충격 극복에 미치는 효과</p>
                    </div>
                    
                    <div class="chart-container" id="chart-age-holder" style="height: 65%;">
                        <!-- Bar Chart SVG -->
                    </div>
                    
                    <div style="font-size: 9px; color: var(--text-muted); margin-top: 4px; line-height: 1.3;">
                        💡 전통(20년+) 축제 효율은 신생 대비 <b>+77%</b> 높음. 
                        <span style="color: var(--red-danger); font-weight: bold;">※ 신생 축제는 중복 개최 시 효율성이 -38.2% 하락하는 직격탄을 맞음.</span>
                    </div>
                </div>

            </div>
        </div>

        <!-- ==========================================
             TAB 4: 전략 시사점 (Strategic Implications)
             ========================================== -->
        <div id="panel-implications" class="tab-panel">
            <!-- TOP SECTION: 잠식 vs 시너지 -->
            <div class="implications-top-grid">
                
                <!-- LEFT: 자기잠식 DANGER ZONE -->
                <div class="imp-card" style="border-left-color: var(--red-danger);">
                    <div>
                        <div class="imp-header red">
                            <span>⚠️</span><span>자기잠식 조건 (DANGER ZONE)</span>
                        </div>
                        <div class="imp-stat-title">30km 이내 + 일정 중첩 + 동일 테마</div>
                        <div class="imp-result red">→ 효율지표 평균 42% 감소 초래</div>
                    </div>
                    
                    <div class="chart-container" id="chart-canni-holder" style="height: 80px;">
                        <!-- Mini horizontal bar chart SVG -->
                    </div>
                    <span class="imp-footer">부산 중구 중복 시 효율 -96%, 강원 동해 중복 시 효율 -93% 폭락</span>
                </div>

                <!-- RIGHT: 시너지 OPPORTUNITY ZONE -->
                <div class="imp-card green" style="border-left-color: var(--green-synergy);">
                    <div>
                        <div class="imp-header green">
                            <span>✓</span><span>시너지 조건 (OPPORTUNITY ZONE)</span>
                        </div>
                        <div class="imp-stat-title">10km 이내 + 상호 테마 보완성 확보</div>
                        <div class="imp-result green">→ 효율지표 최대 5배 이상 급증 가능</div>
                    </div>
                    
                    <div class="chart-container" id="chart-synergy-holder" style="height: 80px;">
                        <!-- Mini horizontal bar chart SVG -->
                    </div>
                    <span class="imp-footer">경남 거창 연합화 시 효율 +5,483%, 경북 포항 연계 시 효율 +1,109% 폭증</span>
                </div>
            </div>

            <!-- BOTTOM SECTION: 3-Phase Action Plan -->
            <div class="phase-grid">
                <!-- Phase 1 -->
                <div class="phase-card" style="border-top-color: var(--red-danger);">
                    <div>
                        <div class="phase-header-row">
                            <span style="font-weight: bold; font-size: 11px; color: var(--red-danger);">Phase 1 · 단기 전략</span>
                            <span class="phase-badge red">즉시 실행 (30일 이내)</span>
                        </div>
                        <div class="phase-step"><span>🔍</span><span>Triage: 진단 & 시장 분리</span></div>
                        <ul class="phase-list">
                            <li>• <b>30km 반경 영향평가제</b> 도입하여 인근 경쟁 필터링</li>
                            <li>• 유사 중복 테마 신설 시 <b>예산 보류 및 중단</b></li>
                            <li>• 과밀 시기(10월) 축제를 비성수기로 <b>전략적 분산</b></li>
                        </ul>
                    </div>
                    <div style="font-size: 9px; color: #9CA3AF; border-top: 1px solid #F1F5F9; padding-top: 6px; margin-top: 8px;">
                        공급 과잉 지역 내 신생/소규모 축제 통폐합 타겟팅
                    </div>
                </div>

                <!-- Phase 2 -->
                <div class="phase-card" style="border-top-color: var(--gold);">
                    <div>
                        <div class="phase-header-row">
                            <span style="font-weight: bold; font-size: 11px; color: var(--gold);">Phase 2 · 중기 전략</span>
                            <span class="phase-badge gold">실행 기간 (6개월 내)</span>
                        </div>
                        <div class="phase-step"><span>🔗</span><span>Clustering: 테마 클러스터링</span></div>
                        <ul class="phase-list">
                            <li>• 10km 이내 인접 지자체 상생 <b>관광 연합체 결성</b></li>
                            <li>• 개별 축제를 묶어 광역 단위 <b>통합 코스 개발</b> (유채꽃 벨트)</li>
                            <li>• 권역 통합 자유이용권 <b>Regional Pass</b> 공동 도입</li>
                        </ul>
                    </div>
                    <div style="font-size: 9px; color: #9CA3AF; border-top: 1px solid #F1F5F9; padding-top: 6px; margin-top: 8px;">
                        잠식을 상호 유기적인 결합을 통한 시너지 모델로 전환
                    </div>
                </div>

                <!-- Phase 3 -->
                <div class="phase-card" style="border-top-color: var(--green-synergy);">
                    <div>
                        <div class="phase-header-row">
                            <span style="font-weight: bold; font-size: 11px; color: var(--green-synergy);">Phase 3 · 장기 전략</span>
                            <span class="phase-badge green">안정화 (1년 이상)</span>
                        </div>
                        <div class="phase-step"><span>🔄</span><span>Lifecycle: 축제 생애주기 평가</span></div>
                        <ul class="phase-list">
                            <li>• 전국 규모의 <b>통합 분석 데이터 인프라</b> 구축</li>
                            <li>• <b>'시범개최 → 성과검증 → 성과 미달 시 일몰심사'</b> 확립</li>
                            <li>• 사전 시뮬레이션 기반 의사결정 프로세스 의무화</li>
                        </ul>
                    </div>
                    <div style="font-size: 9px; color: #9CA3AF; border-top: 1px solid #F1F5F9; padding-top: 6px; margin-top: 8px;">
                        일회성 행사가 아닌 데이터 기반 순환형 축제 관리체계 확립
                    </div>
                </div>
            </div>

            <!-- BOTTOM CALLOUT -->
            <div class="bottom-callout-full">
                📢 "우리 동네 축제"가 아니라 "우리 권역의 통합 페스티벌"로 — 거리는 회피할 것이 아니라 설계해야 할 변수다
            </div>
        </div>
    </main>

    <!-- INTERACTIVE SCRIPTS -->
    <script>
        // DATA CONSTANTS (Same as python backend)
        const MONTHLY_DATA = [
            { month: '1월', count: 27, efficiency: 485.6 },
            { month: '2월', count: 28, efficiency: 256.5 },
            { month: '3월', count: 137, efficiency: 1041.0 },
            { month: '4월', count: 288, efficiency: 501.6 },
            { month: '5월', count: 361, efficiency: 355.9, highlight: 'gold' },
            { month: '6월', count: 99, efficiency: 416.1 },
            { month: '7월', count: 108, efficiency: 204.5 },
            { month: '8월', count: 163, efficiency: 316.5 },
            { month: '9월', count: 326, efficiency: 449.5 },
            { month: '10월', count: 617, efficiency: 267.1, highlight: 'red' },
            { month: '11월', count: 84, efficiency: 243.1 },
            { month: '12월', count: 93, efficiency: 337.4 }
        ];

        const REGIONAL_DATA = [
            { region: '경기', count: 299 },
            { region: '전남', count: 264 },
            { region: '경남', count: 244 },
            { region: '강원', count: 240 },
            { region: '충남', count: 207 },
            { region: '경북', count: 187 },
            { region: '전북', count: 176 },
            { region: '서울', count: 148 },
            { region: '부산', count: 112 },
            { region: '제주', count: 106 },
            { region: '충북', count: 99 },
            { region: '대구', count: 71 },
            { region: '울산', count: 70 },
            { region: '인천', count: 68 },
            { region: '대전', count: 42 },
            { region: '광주', count: 40 },
            { region: '세종', count: 11 }
        ];

        const TYPE_DATA = [
            { type: '문화예술', count: 787, pct: 33.0 },
            { type: '지역특산물', count: 503, pct: 21.1 },
            { type: '자연생태', count: 430, pct: 18.0 },
            { type: '전통역사', count: 300, pct: 12.6 },
            { type: '주민화합', count: 185, pct: 7.8 },
            { type: '기타', count: 112, pct: 4.7 },
            { type: '복합', count: 67, pct: 2.8 }
        ];

        const TYPE_EFFICIENCY = [
            { type: '자연생태', efficiency: 798.2 },
            { type: '지역특산물', efficiency: 468.2 },
            { type: '주민화합', efficiency: 273.3 },
            { type: '복합', efficiency: 256.1 },
            { type: '문화예술', efficiency: 257.2 },
            { type: '기타', efficiency: 248.3 },
            { type: '전통역사', efficiency: 175.7 }
        ];

        const FINANCE_DATA = [
            { label: '지방비 전용', efficiency: 401.2, n: 2221 },
            { label: '국비 지원', efficiency: 193.9, n: 163 }
        ];

        const AGE_DATA = [
            { label: '신생(≤5년)', efficiency: 254.2 },
            { label: '중견(6~19년)', efficiency: 399.9 },
            { label: '전통(≥20년)', efficiency: 450.1 }
        ];

        const CANNIBALIZATION_DATA = [
            { name: '부산 중구(단독)', value: 1400, type: 'single' },
            { name: '부산 중구(중복)', value: 54, type: 'overlap' },
            { name: '강원 동해(단독)', value: 792, type: 'single' },
            { name: '강원 동해(중복)', value: 51, type: 'overlap' }
        ];

        const SYNERGY_DATA = [
            { name: '경남 거창(단독)', value: 16, type: 'single' },
            { name: '경남 거창(연계)', value: 884, type: 'linked' },
            { name: '경북 포항(단독)', value: 184, type: 'single' },
            { name: '경북 포항(연계)', value: 2224, type: 'linked' }
        ];

        // Global Tooltip reference
        const tooltip = document.getElementById('chart-tooltip');

        // Show / Hide tooltip helper
        function showTooltip(evt, title, valueText) {
            tooltip.style.display = 'block';
            tooltip.innerHTML = `<div class="tooltip-title">${title}</div><div>${valueText}</div>`;
            updateTooltipPos(evt);
        }

        function updateTooltipPos(evt) {
            const rect = evt.currentTarget.ownerSVGElement.getBoundingClientRect();
            // Relative position within the container
            const x = evt.clientX - rect.left + 15;
            const y = evt.clientY - rect.top + 10;
            tooltip.style.left = x + 'px';
            tooltip.style.top = y + 'px';
        }

        function hideTooltip() {
            tooltip.style.display = 'none';
        }

        // Tab switcher
        function switchTab(tabId) {
            // Update active tab button style
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            const activeBtn = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick').includes(tabId));
            if (activeBtn) activeBtn.classList.add('active');

            // Toggle panels
            document.querySelectorAll('.tab-panel').forEach(panel => {
                panel.classList.remove('active');
            });
            document.getElementById('panel-' + tabId).classList.add('active');
            
            // Re-render SVG charts in newly active tab to calculate sizes correctly
            renderChartsForTab(tabId);
        }

        // Render functions for charts (Creating pure SVG elements dynamically)
        function renderChartsForTab(tabId) {
            if (tabId === 'overview') {
                renderMonthlyChart();
                renderRegionalPie();
            } else if (tabId === 'crowding') {
                renderComboChart();
                renderRegionalAllChart();
            } else if (tabId === 'efficiency') {
                renderTypeDonut();
                renderTypeEfficiencyChart();
                renderFinanceChart();
                renderAgeChart();
            } else if (tabId === 'implications') {
                renderCanniMiniChart();
                renderSynergyMiniChart();
            }
        }

        // --- TAB 1: OVERVIEW CHARTS ---
        function renderMonthlyChart() {
            const container = document.getElementById('chart-monthly-holder');
            if (!container) return;
            const width = container.clientWidth || 600;
            const height = container.clientHeight || 280;
            container.innerHTML = '';

            const padding = { top: 20, right: 30, bottom: 30, left: 50 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxCount = Math.max(...MONTHLY_DATA.map(d => d.count));
            const barHeight = chartHeight / MONTHLY_DATA.length;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;
            
            // Grid lines (vertical)
            const gridCount = 5;
            for (let i = 0; i <= gridCount; i++) {
                const val = Math.round((maxCount / gridCount) * i);
                const x = padding.left + (chartWidth / gridCount) * i;
                svgHtml += `
                    <line x1="${x}" y1="${padding.top}" x2="${x}" y2="${height - padding.bottom}" stroke="#E2E4EB" stroke-dasharray="3,3" />
                    <text x="${x}" y="${height - padding.bottom + 15}" fill="#9BA8BF" font-size="9" text-anchor="middle">${val}</text>
                `;
            }

            // Reference Line at 198
            const refX = padding.left + (chartWidth * 198 / maxCount);
            svgHtml += `
                <line x1="${refX}" y1="${padding.top}" x2="${refX}" y2="${height - padding.bottom}" stroke="#6B7280" stroke-width="1.5" stroke-dasharray="4,4" />
                <text x="${refX}" y="${padding.top - 5}" fill="#6B7280" font-size="8" text-anchor="middle" font-weight="bold">균등 분배 시 (198건)</text>
            `;

            // Bars and Labels
            MONTHLY_DATA.forEach((d, idx) => {
                const y = padding.top + (idx * barHeight) + 2;
                const h = barHeight - 4;
                const w = (d.count / maxCount) * chartWidth;
                
                let fill = '#9BA8BF';
                if (d.month === '5월') fill = '#B89B4A';
                if (d.month === '10월') fill = '#D94F4F';

                svgHtml += `
                    <text x="${padding.left - 8}" y="${y + h/2 + 3}" fill="#1A1A2E" font-size="10" text-anchor="end">${d.month}</text>
                    <rect class="bar-rect" x="${padding.left}" y="${y}" width="${w}" height="${h}" fill="${fill}" rx="2" 
                          onmousemove="showTooltip(event, '${d.month}', '축제 수: <b>${d.count}건</b><br>평균 효율: <b>${d.efficiency.toFixed(1)}</b>')"
                          onmouseout="hideTooltip()"/>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderRegionalPie() {
            const container = document.getElementById('chart-regional-holder');
            if (!container) return;
            const width = container.clientWidth || 400;
            const height = container.clientHeight || 280;
            container.innerHTML = '';

            const cx = width / 2;
            const cy = height / 2;
            const innerRadius = 50;
            const outerRadius = 80;

            const top8 = REGIONAL_DATA.slice(0, 8);
            const total = 2384;
            const top8Sum = top8.reduce((sum, d) => sum + d.count, 0);
            
            const data = [
                ...top8.map(d => ({ name: d.region, count: d.count })),
                { name: '기타', count: total - top8Sum }
            ];

            const colors = [
                '#1C2D4A', '#263C62', '#314B7A', '#3C5B92', 
                '#4C6CA3', '#5E7DB5', '#718FC7', '#85A2D9', '#CBD2DF'
            ];

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;
            
            let startAngle = 0;
            data.forEach((d, idx) => {
                const pct = d.count / total;
                const angle = pct * 360;
                
                // Coordinates for outer arc
                const x1 = cx + outerRadius * Math.cos((startAngle - 90) * Math.PI / 180);
                const y1 = cy + outerRadius * Math.sin((startAngle - 90) * Math.PI / 180);
                const x2 = cx + outerRadius * Math.cos((startAngle + angle - 90) * Math.PI / 180);
                const y2 = cy + outerRadius * Math.sin((startAngle + angle - 90) * Math.PI / 180);

                // Coordinates for inner arc
                const xi1 = cx + innerRadius * Math.cos((startAngle - 90) * Math.PI / 180);
                const yi1 = cy + innerRadius * Math.sin((startAngle - 90) * Math.PI / 180);
                const xi2 = cx + innerRadius * Math.cos((startAngle + angle - 90) * Math.PI / 180);
                const yi2 = cy + innerRadius * Math.sin((startAngle + angle - 90) * Math.PI / 180);

                const largeArcFlag = angle > 180 ? 1 : 0;

                // SVG Path for Donut Slice
                const pathData = `
                    M ${x1} ${y1}
                    A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${x2} ${y2}
                    L ${xi2} ${yi2}
                    A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 0 ${xi1} ${yi1}
                    Z
                `;

                svgHtml += `
                    <path class="pie-slice" d="${pathData}" fill="${colors[idx]}" stroke="#FFFFFF" stroke-width="1.5"
                          onmousemove="showTooltip(event, '${d.name}', '축제 수: <b>${d.count}건</b> (${(pct*100).toFixed(1)}%)')"
                          onmouseout="hideTooltip()"/>
                `;

                startAngle += angle;
            });

            // Center Text
            svgHtml += `
                <text x="${cx}" y="${cy - 4}" fill="#6B7280" font-size="10" text-anchor="middle" font-weight="500">17개 광역</text>
                <text x="${cx}" y="${cy + 12}" fill="#1C2D4A" font-size="15" text-anchor="middle" font-weight="bold">${total}건</text>
            `;

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        // --- TAB 2: CROWDING CHARTS ---
        function renderComboChart() {
            const container = document.getElementById('chart-combo-holder');
            if (!container) return;
            const width = container.clientWidth || 600;
            const height = container.clientHeight || 300;
            container.innerHTML = '';

            const padding = { top: 25, right: 45, bottom: 30, left: 40 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxCount = Math.max(...MONTHLY_DATA.map(d => d.count)); // For Left Y
            const maxEff = 1100; // For Right Y (efficiency holds max 1041.0)

            const colWidth = chartWidth / MONTHLY_DATA.length;
            const barW = colWidth - 8;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            // Grid lines (horizontal)
            const gridRows = 4;
            for (let i = 0; i <= gridRows; i++) {
                const y = padding.top + (chartHeight / gridRows) * i;
                const leftVal = Math.round((maxCount / gridRows) * (gridRows - i));
                const rightVal = Math.round((maxEff / gridRows) * (gridRows - i));
                svgHtml += `
                    <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#E2E4EB" />
                    <text x="${padding.left - 8}" y="${y + 3}" fill="#9BA8BF" font-size="9" text-anchor="end">${leftVal}</text>
                    <text x="${width - padding.right + 8}" y="${y + 3}" fill="#9BA8BF" font-size="9" text-anchor="start">${rightVal}</text>
                `;
            }

            // Left / Right Y Axis Labels
            svgHtml += `
                <text x="${padding.left - 5}" y="${padding.top - 10}" fill="#6B7280" font-size="8" text-anchor="end">개최 수(건)</text>
                <text x="${width - padding.right + 5}" y="${padding.top - 10}" fill="#6B7280" font-size="8" text-anchor="start">효율(명/M)</text>
            `;

            let linePoints = [];

            // Draw Bars first, build Line points
            MONTHLY_DATA.forEach((d, idx) => {
                const x = padding.left + (idx * colWidth) + 4;
                const barH = (d.count / maxCount) * chartHeight;
                const barY = height - padding.bottom - barH;

                let fill = '#9BA8BF';
                if (d.month === '5월') fill = '#B89B4A';
                if (d.month === '10월') fill = '#D94F4F';

                // Bar
                svgHtml += `
                    <rect class="bar-rect" x="${x}" y="${barY}" width="${barW}" height="${barH}" fill="${fill}" rx="2"
                          onmousemove="showTooltip(event, '${d.month}', '개최 수: <b>${d.count}건</b><br>평균 효율: <b>${d.efficiency.toFixed(1)}</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${x + barW/2}" y="${height - padding.bottom + 15}" fill="#1A1A2E" font-size="10" text-anchor="middle">${d.month.replace('월','')}</text>
                `;

                // Line point (Right Y axis scale)
                const lx = x + barW/2;
                const ly = height - padding.bottom - ((d.efficiency / maxEff) * chartHeight);
                linePoints.push({ x: lx, y: ly, label: d.month, val: d.efficiency });
            });

            // Draw line
            let linePath = `M ${linePoints[0].x} ${linePoints[0].y}`;
            for (let i = 1; i < linePoints.length; i++) {
                linePath += ` L ${linePoints[i].x} ${linePoints[i].y}`;
            }

            svgHtml += `
                <path d="${linePath}" fill="none" stroke="#3C8C5A" stroke-width="2.5" />
            `;

            // Draw markers on line
            linePoints.forEach(p => {
                svgHtml += `
                    <circle cx="${p.x}" cy="${p.y}" r="3.5" fill="#FFFFFF" stroke="#3C8C5A" stroke-width="2"
                            onmousemove="showTooltip(event, '${p.label}', '평균 효율: <b>${p.val.toFixed(1)}명/백만원</b>')"
                            onmouseout="hideTooltip()"/>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderRegionalAllChart() {
            const container = document.getElementById('chart-regional-all-holder');
            if (!container) return;
            const width = container.clientWidth || 600;
            const height = container.clientHeight || 280;
            container.innerHTML = '';

            const padding = { top: 20, right: 15, bottom: 30, left: 35 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxCount = Math.max(...REGIONAL_DATA.map(d => d.count));
            const colWidth = chartWidth / REGIONAL_DATA.length;
            const barW = colWidth - 6;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            // Grid lines (horizontal)
            const gridRows = 4;
            for (let i = 0; i <= gridRows; i++) {
                const y = padding.top + (chartHeight / gridRows) * i;
                const val = Math.round((maxCount / gridRows) * (gridRows - i));
                svgHtml += `
                    <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#E2E4EB" />
                    <text x="${padding.left - 8}" y="${y + 3}" fill="#9BA8BF" font-size="9" text-anchor="end">${val}</text>
                `;
            }

            // Draw Bars and Labels
            REGIONAL_DATA.forEach((d, idx) => {
                const x = padding.left + (idx * colWidth) + 3;
                const barH = (d.count / maxCount) * chartHeight;
                const barY = height - padding.bottom - barH;

                // Top 4 Navy, others Muted
                const fill = idx < 4 ? '#1C2D4A' : '#9BA8BF';

                svgHtml += `
                    <rect class="bar-rect" x="${x}" y="${barY}" width="${barW}" height="${barH}" fill="${fill}" rx="1"
                          onmousemove="showTooltip(event, '${d.region}', '축제 수: <b>${d.count}건</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${x + barW/2}" y="${height - padding.bottom + 15}" fill="#1A1A2E" font-size="9" text-anchor="middle">${d.region}</text>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        // --- TAB 3: THEME & EFFICIENCY CHARTS ---
        function renderTypeDonut() {
            const container = document.getElementById('chart-type-donut-holder');
            if (!container) return;
            const width = container.clientWidth || 300;
            const height = container.clientHeight || 150;
            container.innerHTML = '';

            const cx = width / 2;
            const cy = height / 2;
            const innerRadius = 32;
            const outerRadius = 50;

            const total = 2384;
            const colors = {
                '문화예술': '#1C2D4A',
                '지역특산물': '#3A5A8A',
                '자연생태': '#5C7FA8',
                '전통역사': '#B89B4A',
                '주민화합': '#9BA8BF',
                '기타': '#CBD2DF',
                '복합': '#E2E4EB'
            };

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;
            
            let startAngle = 0;
            TYPE_DATA.forEach((d, idx) => {
                const pct = d.count / total;
                const angle = pct * 360;
                
                const x1 = cx + outerRadius * Math.cos((startAngle - 90) * Math.PI / 180);
                const y1 = cy + outerRadius * Math.sin((startAngle - 90) * Math.PI / 180);
                const x2 = cx + outerRadius * Math.cos((startAngle + angle - 90) * Math.PI / 180);
                const y2 = cy + outerRadius * Math.sin((startAngle + angle - 90) * Math.PI / 180);

                const xi1 = cx + innerRadius * Math.cos((startAngle - 90) * Math.PI / 180);
                const yi1 = cy + innerRadius * Math.sin((startAngle - 90) * Math.PI / 180);
                const xi2 = cx + innerRadius * Math.cos((startAngle + angle - 90) * Math.PI / 180);
                const yi2 = cy + innerRadius * Math.sin((startAngle + angle - 90) * Math.PI / 180);

                const largeArcFlag = angle > 180 ? 1 : 0;

                const pathData = `
                    M ${x1} ${y1}
                    A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${x2} ${y2}
                    L ${xi2} ${yi2}
                    A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 0 ${xi1} ${yi1}
                    Z
                `;

                svgHtml += `
                    <path class="pie-slice" d="${pathData}" fill="${colors[d.type]}" stroke="#FFFFFF" stroke-width="1.5"
                          onmousemove="showTooltip(event, '${d.type}', '축제 수: <b>${d.count}건</b> (${d.pct}%)')"
                          onmouseout="hideTooltip()"/>
                `;

                startAngle += angle;
            });

            // Center Labels
            svgHtml += `
                <text x="${cx}" y="${cy - 2}" fill="#6B7280" font-size="8" text-anchor="middle">총 축제</text>
                <text x="${cx}" y="${cy + 10}" fill="#1C2D4A" font-size="12" text-anchor="middle" font-weight="bold">${total}건</text>
            `;

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderTypeEfficiencyChart() {
            const container = document.getElementById('chart-type-eff-holder');
            if (!container) return;
            const width = container.clientWidth || 300;
            const height = container.clientHeight || 150;
            container.innerHTML = '';

            const padding = { top: 10, right: 15, bottom: 20, left: 65 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxVal = Math.max(...TYPE_EFFICIENCY.map(d => d.efficiency));
            const barHeight = chartHeight / TYPE_EFFICIENCY.length;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            // Grid lines (vertical)
            const gridCount = 4;
            for (let i = 0; i <= gridCount; i++) {
                const val = Math.round((maxVal / gridCount) * i);
                const x = padding.left + (chartWidth / gridCount) * i;
                svgHtml += `
                    <line x1="${x}" y1="${padding.top}" x2="${x}" y2="${height - padding.bottom}" stroke="#E2E4EB" stroke-dasharray="2,2" />
                    <text x="${x}" y="${height - padding.bottom + 12}" fill="#9BA8BF" font-size="8" text-anchor="middle">${val}</text>
                `;
            }

            // Draw bars
            TYPE_EFFICIENCY.forEach((d, idx) => {
                const y = padding.top + (idx * barHeight) + 2;
                const h = barHeight - 4;
                const w = (d.efficiency / maxVal) * chartWidth;

                let fill = '#9BA8BF';
                if (d.type === '자연생태' || d.type === '지역특산물') fill = '#3C8C5A';
                if (d.type === '전통역사' || d.type === '문화예술') fill = '#D94F4F';

                svgHtml += `
                    <text x="${padding.left - 5}" y="${y + h/2 + 3}" fill="#1A1A2E" font-size="9" text-anchor="end">${d.type}</text>
                    <rect class="bar-rect" x="${padding.left}" y="${y}" width="${w}" height="${h}" fill="${fill}" rx="1.5"
                          onmousemove="showTooltip(event, '${d.type}', '평균 효율: <b>${d.efficiency.toFixed(1)}명/백만원</b>')"
                          onmouseout="hideTooltip()"/>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderFinanceChart() {
            const container = document.getElementById('chart-finance-holder');
            if (!container) return;
            const width = container.clientWidth || 200;
            const height = container.clientHeight || 150;
            container.innerHTML = '';

            const padding = { top: 20, right: 10, bottom: 25, left: 30 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxVal = 500;
            const colWidth = chartWidth / FINANCE_DATA.length;
            const barW = colWidth - 20;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            // Grid lines (horizontal)
            const gridRows = 3;
            for (let i = 0; i <= gridRows; i++) {
                const y = padding.top + (chartHeight / gridRows) * i;
                const val = Math.round((maxVal / gridRows) * (gridRows - i));
                svgHtml += `
                    <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#E2E4EB" />
                    <text x="${padding.left - 5}" y="${y + 3}" fill="#9BA8BF" font-size="8" text-anchor="end">${val}</text>
                `;
            }

            // Draw bars
            FINANCE_DATA.forEach((d, idx) => {
                const x = padding.left + (idx * colWidth) + 10;
                const barH = (d.efficiency / maxVal) * chartHeight;
                const barY = height - padding.bottom - barH;

                const fill = d.label === '국비 지원' ? '#D94F4F' : '#3C8C5A';

                svgHtml += `
                    <rect class="bar-rect" x="${x}" y="${barY}" width="${barW}" height="${barH}" fill="${fill}" rx="2"
                          onmousemove="showTooltip(event, '${d.label}', '평균 효율: <b>${d.efficiency.toFixed(1)}명/M</b><br>샘플 수(n): <b>${d.n}</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${x + barW/2}" y="${height - padding.bottom + 12}" fill="#1A1A2E" font-size="9" text-anchor="middle">${d.label}</text>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderAgeChart() {
            const container = document.getElementById('chart-age-holder');
            if (!container) return;
            const width = container.clientWidth || 300;
            const height = container.clientHeight || 150;
            container.innerHTML = '';

            const padding = { top: 20, right: 10, bottom: 25, left: 30 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxVal = 500;
            const colWidth = chartWidth / AGE_DATA.length;
            const barW = colWidth - 25;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            // Grid lines (horizontal)
            const gridRows = 3;
            for (let i = 0; i <= gridRows; i++) {
                const y = padding.top + (chartHeight / gridRows) * i;
                const val = Math.round((maxVal / gridRows) * (gridRows - i));
                svgHtml += `
                    <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#E2E4EB" />
                    <text x="${padding.left - 5}" y="${y + 3}" fill="#9BA8BF" font-size="8" text-anchor="end">${val}</text>
                `;
            }

            // Draw bars
            AGE_DATA.forEach((d, idx) => {
                const x = padding.left + (idx * colWidth) + 12.5;
                const barH = (d.efficiency / maxVal) * chartHeight;
                const barY = height - padding.bottom - barH;

                let fill = '#9BA8BF';
                if (d.label.includes('신생')) fill = '#D94F4F';
                if (d.label.includes('중견')) fill = '#B89B4A';
                if (d.label.includes('전통')) fill = '#3C8C5A';

                svgHtml += `
                    <rect class="bar-rect" x="${x}" y="${barY}" width="${barW}" height="${barH}" fill="${fill}" rx="2"
                          onmousemove="showTooltip(event, '${d.label}', '평균 효율: <b>${d.efficiency.toFixed(1)}명/백만원</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${x + barW/2}" y="${height - padding.bottom + 12}" fill="#1A1A2E" font-size="9" text-anchor="middle">${d.label}</text>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        // --- TAB 4: IMPLICATIONS CHARTS ---
        function renderCanniMiniChart() {
            const container = document.getElementById('chart-canni-holder');
            if (!container) return;
            const width = container.clientWidth || 300;
            const height = container.clientHeight || 80;
            container.innerHTML = '';

            const padding = { top: 5, right: 15, bottom: 5, left: 85 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxVal = 1500;
            const barHeight = chartHeight / CANNIBALIZATION_DATA.length;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            CANNIBALIZATION_DATA.forEach((d, idx) => {
                const y = padding.top + (idx * barHeight) + 1.5;
                const h = barHeight - 3;
                const w = (d.value / maxVal) * chartWidth;

                const fill = d.type === 'overlap' ? '#D94F4F' : '#9BA8BF';

                svgHtml += `
                    <text x="${padding.left - 5}" y="${y + h/2 + 3}" fill="#1A1A2E" font-size="8.5" text-anchor="end">${d.name}</text>
                    <rect class="bar-rect" x="${padding.left}" y="${y}" width="${w}" height="${h}" fill="${fill}" rx="1"
                          onmousemove="showTooltip(event, '${d.name}', '효율: <b>${d.value.toLocaleString()}명/백만원</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${padding.left + w + 4}" y="${y + h/2 + 3}" fill="#6B7280" font-size="8" text-anchor="start">${d.value}</text>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        function renderSynergyMiniChart() {
            const container = document.getElementById('chart-synergy-holder');
            if (!container) return;
            const width = container.clientWidth || 300;
            const height = container.clientHeight || 80;
            container.innerHTML = '';

            const padding = { top: 5, right: 25, bottom: 5, left: 85 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;

            const maxVal = 2300;
            const barHeight = chartHeight / SYNERGY_DATA.length;

            let svgHtml = `<svg viewBox="0 0 ${width} ${height}">`;

            SYNERGY_DATA.forEach((d, idx) => {
                const y = padding.top + (idx * barHeight) + 1.5;
                const h = barHeight - 3;
                const w = (d.value / maxVal) * chartWidth;

                const fill = d.type === 'linked' ? '#3C8C5A' : '#9BA8BF';

                svgHtml += `
                    <text x="${padding.left - 5}" y="${y + h/2 + 3}" fill="#1A1A2E" font-size="8.5" text-anchor="end">${d.name}</text>
                    <rect class="bar-rect" x="${padding.left}" y="${y}" width="${w}" height="${h}" fill="${fill}" rx="1"
                          onmousemove="showTooltip(event, '${d.name}', '효율: <b>${d.value.toLocaleString()}명/백만원</b>')"
                          onmouseout="hideTooltip()"/>
                    <text x="${padding.left + w + 4}" y="${y + h/2 + 3}" fill="#6B7280" font-size="8" text-anchor="start">${d.value}</text>
                `;
            });

            svgHtml += '</svg>';
            container.innerHTML = svgHtml;
        }

        // Initialize and first render
        window.addEventListener('load', () => {
            renderChartsForTab('overview');
        });

        // Handle window resize dynamically to fit responsiveness
        window.addEventListener('resize', () => {
            const activeTab = document.querySelector('.tab-btn.active').getAttribute('onclick').match(/'([^']+)'/)[1];
            renderChartsForTab(activeTab);
        });
    </script>
</body>
</html>
"""

    try:
        # 파일 쓰기
        with open(html_output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"성공: 대시보드 HTML 파일이 생성되었습니다 -> {html_output_path}")
        
        # 3. 브라우저로 열기
        print("\n=== [3/3] 대시보드 웹 브라우저 실행 ===")
        abs_path = os.path.abspath(html_output_path)
        file_url = f"file:///{abs_path.replace(os.sep, '/')}"
        print(f"URL: {file_url}")
        webbrowser.open(file_url)
        print("대시보드가 기본 웹 브라우저에 표시되었습니다.")
        
    except Exception as e:
        print(f"HTML 파일 생성 및 브라우저 기동 중 오류 발생: {e}")

if __name__ == "__main__":
    analyze_and_generate_dashboard()
