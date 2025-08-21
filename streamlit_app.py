import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

from redmine_client import RedmineClient
from ppt_generator import PowerPointGenerator

st.set_page_config(
    page_title="Redmineチケット可視化ダッシュボード",
    page_icon="📊",
    layout="wide"
)

def show_welcome_screen():
    """ウェルカム画面を表示"""
    st.markdown("""
    <div style='text-align: center; padding: 2rem;'>
        <h1>🎯 Redmineチケット可視化ダッシュボード</h1>
        <p style='font-size: 1.2em; color: #666;'>
            RedmineのAPIを使用してチケットデータを取得し、<br>
            インタラクティブなグラフとレポートで可視化します
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### 🔧 接続設定")
        
        # RedmineサーバーURL入力
        redmine_url = st.text_input(
            "RedmineサーバーURL",
            value="http://localhost:3000",
            placeholder="例: https://your-redmine-server.com",
            help="RedmineサーバーのベースURLを入力してください"
        )
        
        # APIキー入力
        api_key = st.text_input(
            "APIキー",
            type="password",
            placeholder="Redmine APIキーを入力してください",
            help="Redmine管理画面の「個人設定」> 「APIアクセスキー」で確認できます"
        )
        
        st.markdown("---")
        
        # 接続ボタン
        if st.button("🚀 ダッシュボードを開始", type="primary", use_container_width=True):
            if not redmine_url.strip():
                st.error("RedmineサーバーURLを入力してください")
                return False
            if not api_key.strip():
                st.error("APIキーを入力してください")
                return False
            
            # 接続テスト
            with st.spinner("Redmine接続を確認中..."):
                try:
                    test_client = RedmineClient(redmine_url.strip(), api_key.strip())
                    # 接続テスト（少量のデータ取得）
                    test_data = test_client.get_issues(limit=1)
                    
                    # セッション状態に保存
                    st.session_state.redmine_url = redmine_url.strip()
                    st.session_state.api_key = api_key.strip()
                    st.session_state.connected = True
                    
                    st.success("✅ 接続成功！ダッシュボードに移動します...")
                    st.rerun()
                    return True
                    
                except Exception as e:
                    st.error(f"❌ 接続エラー: {str(e)}")
                    st.info("以下を確認してください：\n- RedmineサーバーURLが正しいか\n- APIキーが有効か\n- RedmineのAPI機能が有効になっているか")
                    return False
        
        st.markdown("---")
        
        # ヘルプセクション
        with st.expander("💡 APIキーの取得方法"):
            st.markdown("""
            1. Redmineにログインします
            2. 右上のユーザー名をクリック → **個人設定**
            3. 右側メニューの **APIアクセスキー** をクリック
            4. **表示** ボタンをクリックしてAPIキーをコピー
            5. 上記のフィールドにペーストしてください
            """)
        
        with st.expander("🔍 機能紹介"):
            st.markdown("""
            **📊 可視化機能**
            - ステータス・優先度・担当者別の分析グラフ
            - スケジュール・期限分析とガントチャート
            - 工数分析と進捗管理
            
            **⚠️ アラート機能**
            - 期限超過チケットの自動検出
            - 期限間近チケットの事前通知
            
            **📄 エクスポート機能**
            - チケット一覧のCSVダウンロード
            - 個別チケットのPowerPoint帳票出力
            """)
    
    return False

@st.cache_data
def load_redmine_data(_redmine_url, _api_key):
    """Redmineデータを取得（キャッシュ付き）"""
    try:
        client = RedmineClient(_redmine_url, _api_key)
        issues = client.get_all_issues()
        df = client.issues_to_dataframe(issues)
        return df, client
    except Exception as e:
        st.error(f"Redmineからのデータ取得に失敗しました: {e}")
        return pd.DataFrame(), None

def create_status_chart(df):
    if df.empty:
        return None
    
    status_counts = df['ステータス'].value_counts()
    fig = px.pie(
        values=status_counts.values,
        names=status_counts.index,
        title="ステータス別チケット数"
    )
    return fig

def create_priority_chart(df):
    if df.empty:
        return None
    
    priority_counts = df['優先度'].value_counts()
    fig = px.bar(
        x=priority_counts.index,
        y=priority_counts.values,
        title="優先度別チケット数"
    )
    fig.update_layout(xaxis_title="優先度", yaxis_title="チケット数")
    return fig

def create_assignee_chart(df):
    if df.empty:
        return None
    
    assignee_counts = df[df['担当者'] != '']['担当者'].value_counts().head(10)
    fig = px.bar(
        x=assignee_counts.values,
        y=assignee_counts.index,
        orientation='h',
        title="担当者別チケット数 (上位10名)"
    )
    fig.update_layout(xaxis_title="チケット数", yaxis_title="担当者")
    return fig

def create_project_chart(df):
    if df.empty:
        return None
    
    project_counts = df['プロジェクト'].value_counts()
    fig = px.pie(
        values=project_counts.values,
        names=project_counts.index,
        title="プロジェクト別チケット数"
    )
    return fig

def create_tracker_chart(df):
    if df.empty:
        return None
    
    tracker_counts = df['トラッカー'].value_counts()
    fig = px.bar(
        x=tracker_counts.index,
        y=tracker_counts.values,
        title="トラッカー別チケット数"
    )
    fig.update_layout(xaxis_title="トラッカー", yaxis_title="チケット数")
    return fig

def create_deadline_chart(df):
    if df.empty:
        return None
    
    # 期限日が設定されているチケットのみ
    df_with_deadline = df.dropna(subset=['期限日'])
    if df_with_deadline.empty:
        return None
    
    # 期限日を月単位でグループ化
    df_with_deadline['期限月'] = df_with_deadline['期限日'].dt.to_period('M')
    deadline_counts = df_with_deadline.groupby('期限月').size()
    
    fig = px.bar(
        x=[str(period) for period in deadline_counts.index],
        y=deadline_counts.values,
        title="月別期限チケット数",
        color=deadline_counts.values,
        color_continuous_scale='Reds'
    )
    fig.update_layout(
        xaxis_title="期限月", 
        yaxis_title="チケット数",
        showlegend=False
    )
    return fig

def create_schedule_gantt_chart(df):
    if df.empty:
        return None
    
    # 開始日と期限日が両方設定されているチケットのみ
    df_schedule = df.dropna(subset=['開始日', '期限日']).head(20)  # 表示を20件に限定
    if df_schedule.empty:
        return None
    
    fig = px.timeline(
        df_schedule, 
        x_start="開始日", 
        x_end="期限日", 
        y="件名",
        color="ステータス",
        title="チケットスケジュール（上位20件）"
    )
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=600)
    return fig

def create_progress_vs_deadline_chart(df):
    if df.empty:
        return None
    
    # 期限日が設定されているチケット
    df_with_deadline = df.dropna(subset=['期限日'])
    if df_with_deadline.empty:
        return None
    
    # 期限超過の判定
    today = pd.Timestamp.now().normalize()
    df_with_deadline['期限超過'] = (df_with_deadline['期限日'] < today) & (df_with_deadline['進捗率'] < 100)
    df_with_deadline['期限まで日数'] = (df_with_deadline['期限日'] - today).dt.days
    
    fig = px.scatter(
        df_with_deadline,
        x="期限まで日数",
        y="進捗率",
        color="期限超過",
        size="実績工数",
        hover_data=['件名', 'ステータス', '担当者'],
        title="進捗率 vs 期限まで日数",
        color_discrete_map={True: 'red', False: 'blue'}
    )
    fig.update_layout(
        xaxis_title="期限まで日数（負数は超過）",
        yaxis_title="進捗率（%）"
    )
    return fig

def create_workload_chart(df):
    if df.empty:
        return None
    
    # 担当者別の工数集計
    df_assigned = df[df['担当者'] != '']
    if df_assigned.empty:
        return None
    
    workload_data = df_assigned.groupby('担当者').agg({
        '予定工数': 'sum',
        '実績工数': 'sum',
        'ID': 'count'
    }).rename(columns={'ID': 'チケット数'})
    
    workload_data = workload_data.head(10)  # 上位10名
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='予定工数',
        x=workload_data.index,
        y=workload_data['予定工数'],
        yaxis='y',
        offsetgroup=1
    ))
    
    fig.add_trace(go.Bar(
        name='実績工数',
        x=workload_data.index,
        y=workload_data['実績工数'],
        yaxis='y',
        offsetgroup=2
    ))
    
    fig.add_trace(go.Scatter(
        name='チケット数',
        x=workload_data.index,
        y=workload_data['チケット数'],
        yaxis='y2',
        mode='lines+markers',
        line=dict(color='orange', width=3),
        marker=dict(size=8)
    ))
    
    fig.update_layout(
        title='担当者別工数・チケット数',
        xaxis_title='担当者',
        yaxis=dict(
            title='工数（時間）',
            side='left'
        ),
        yaxis2=dict(
            title='チケット数',
            side='right',
            overlaying='y'
        ),
        barmode='group',
        height=400
    )
    
    return fig

def show_dashboard():
    """ダッシュボード画面を表示"""
    # ヘッダー部分
    col1, col2 = st.columns([4, 1])
    with col1:
        st.title("📊 Redmineチケット可視化ダッシュボード")
    with col2:
        if st.button("🔧 設定変更", help="接続設定を変更"):
            # セッション状態をクリアして設定画面に戻る
            for key in ['redmine_url', 'api_key', 'connected']:
                if key in st.session_state:
                    del st.session_state[key]
            # キャッシュもクリア
            st.cache_data.clear()
            st.rerun()
    
    st.markdown("---")
    
    # Redmineデータを取得
    df, client = load_redmine_data(st.session_state.redmine_url, st.session_state.api_key)
    
    if df.empty:
        st.warning("データが取得できませんでした。設定を確認してください。")
        return
    
    st.sidebar.header("フィルター設定")
    
    projects = ['すべて'] + list(df['プロジェクト'].unique())
    selected_project = st.sidebar.selectbox("プロジェクト", projects)
    
    statuses = ['すべて'] + list(df['ステータス'].unique())
    selected_status = st.sidebar.selectbox("ステータス", statuses)
    
    filtered_df = df.copy()
    if selected_project != 'すべて':
        filtered_df = filtered_df[filtered_df['プロジェクト'] == selected_project]
    if selected_status != 'すべて':
        filtered_df = filtered_df[filtered_df['ステータス'] == selected_status]
    
    st.header("📈 概要統計")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("総チケット数", len(filtered_df))
    with col2:
        open_tickets = len(filtered_df[~filtered_df['ステータス'].str.contains('終了|完了|解決済み', na=False)])
        st.metric("未完了チケット数", open_tickets)
    with col3:
        avg_progress = filtered_df['進捗率'].mean()
        st.metric("平均進捗率", f"{avg_progress:.1f}%")
    with col4:
        total_hours = filtered_df['実績工数'].sum()
        st.metric("総実績工数", f"{total_hours:.1f}h")
    
    # アラート表示
    today = pd.Timestamp.now().normalize()
    if not filtered_df.empty:
        # 期限超過チケット
        overdue_tickets = filtered_df[
            (filtered_df['期限日'].notna()) & 
            (filtered_df['期限日'] < today) & 
            (filtered_df['進捗率'] < 100)
        ]
        
        if not overdue_tickets.empty:
            st.error(f"⚠️ 期限超過チケット: {len(overdue_tickets)}件")
            with st.expander("期限超過チケット一覧"):
                display_overdue = overdue_tickets[['ID', '件名', '担当者', '期限日', '進捗率']].copy()
                display_overdue['期限日'] = display_overdue['期限日'].dt.strftime('%Y-%m-%d')
                st.dataframe(display_overdue, use_container_width=True)
        
        # 期限間近チケット（7日以内）
        upcoming_deadline = filtered_df[
            (filtered_df['期限日'].notna()) & 
            (filtered_df['期限日'] >= today) & 
            (filtered_df['期限日'] <= today + pd.Timedelta(days=7)) &
            (filtered_df['進捗率'] < 100)
        ]
        
        if not upcoming_deadline.empty:
            st.warning(f"📅 期限間近チケット（7日以内）: {len(upcoming_deadline)}件")
            with st.expander("期限間近チケット一覧"):
                display_upcoming = upcoming_deadline[['ID', '件名', '担当者', '期限日', '進捗率']].copy()
                display_upcoming['期限日'] = display_upcoming['期限日'].dt.strftime('%Y-%m-%d')
                st.dataframe(display_upcoming, use_container_width=True)
    
    st.markdown("---")
    st.header("📊 チケット分析グラフ")
    
    tab1, tab2, tab3 = st.tabs(["基本分析", "担当者・プロジェクト分析", "スケジュール・期限分析"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            status_chart = create_status_chart(filtered_df)
            if status_chart:
                st.plotly_chart(status_chart, use_container_width=True)
        
        with col2:
            priority_chart = create_priority_chart(filtered_df)
            if priority_chart:
                st.plotly_chart(priority_chart, use_container_width=True)
        
        tracker_chart = create_tracker_chart(filtered_df)
        if tracker_chart:
            st.plotly_chart(tracker_chart, use_container_width=True)
    
    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            assignee_chart = create_assignee_chart(filtered_df)
            if assignee_chart:
                st.plotly_chart(assignee_chart, use_container_width=True)
        
        with col2:
            project_chart = create_project_chart(filtered_df)
            if project_chart:
                st.plotly_chart(project_chart, use_container_width=True)
        
        # 工数分析（フル幅）
        workload_chart = create_workload_chart(filtered_df)
        if workload_chart:
            st.plotly_chart(workload_chart, use_container_width=True)
    
    with tab3:
        st.subheader("📅 スケジュール・期限分析")
        
        col1, col2 = st.columns(2)
        
        with col1:
            deadline_chart = create_deadline_chart(filtered_df)
            if deadline_chart:
                st.plotly_chart(deadline_chart, use_container_width=True)
        
        with col2:
            progress_chart = create_progress_vs_deadline_chart(filtered_df)
            if progress_chart:
                st.plotly_chart(progress_chart, use_container_width=True)
        
        # ガントチャート（フル幅）
        gantt_chart = create_schedule_gantt_chart(filtered_df)
        if gantt_chart:
            st.plotly_chart(gantt_chart, use_container_width=True)
        else:
            st.info("スケジュール表示には開始日と期限日が両方設定されたチケットが必要です。")
    
    st.markdown("---")
    st.header("📄 チケット一覧・詳細・エクスポート")
    
    # タブで機能を分割
    tab1, tab2 = st.tabs(["📋 チケット一覧", "📥 CSVエクスポート"])
    
    with tab1:
        st.subheader("チケット一覧")
        
        # チケット一覧表示
        display_df = filtered_df[['ID', '件名', 'ステータス', '優先度', '担当者', '進捗率', '作成日']].copy()
        if not display_df.empty:
            display_df['作成日'] = display_df['作成日'].dt.strftime('%Y-%m-%d')
            
            # ページネーション設定
            items_per_page = 10
            total_items = len(display_df)
            total_pages = (total_items - 1) // items_per_page + 1
            
            if 'current_page' not in st.session_state:
                st.session_state.current_page = 1
            
            # ページ選択UI
            if total_pages > 1:
                col1, col2, col3 = st.columns([1, 3, 1])
                with col1:
                    if st.button("← 前のページ") and st.session_state.current_page > 1:
                        st.session_state.current_page -= 1
                        st.rerun()
                with col2:
                    st.write(f"ページ {st.session_state.current_page} / {total_pages} ({total_items}件)")
                with col3:
                    if st.button("次のページ →") and st.session_state.current_page < total_pages:
                        st.session_state.current_page += 1
                        st.rerun()
            
            # 現在のページのデータを取得
            start_idx = (st.session_state.current_page - 1) * items_per_page
            end_idx = min(start_idx + items_per_page, total_items)
            page_df = display_df.iloc[start_idx:end_idx].copy()
            
            # チケット選択用のラジオボタン
            if 'selected_ticket_id' not in st.session_state:
                st.session_state.selected_ticket_id = None
            
            # チケット選択UI
            st.markdown("**チケット一覧：**")
            
            # 選択されたチケットIDを追跡
            selected_index = None
            if st.session_state.selected_ticket_id:
                # 現在のページで選択されたチケットのインデックスを見つける
                matching_rows = page_df[page_df['ID'] == st.session_state.selected_ticket_id]
                if not matching_rows.empty:
                    selected_index = matching_rows.index[0] - page_df.index[0]
            
            # ラジオボタンでチケット選択
            ticket_options = []
            for idx, row in page_df.iterrows():
                subject = row['件名'][:50] + "..." if len(row['件名']) > 50 else row['件名']
                option_text = f"#{row['ID']} | {subject} | {row['ステータス']} | {row['担当者'] if row['担当者'] else '未設定'}"
                ticket_options.append((option_text, row['ID']))
            
            if ticket_options:
                # 現在選択されているチケットのインデックスを取得
                current_selection_idx = 0
                if st.session_state.selected_ticket_id:
                    for i, (_, ticket_id) in enumerate(ticket_options):
                        if ticket_id == st.session_state.selected_ticket_id:
                            current_selection_idx = i
                            break
                
                # ラジオボタンで選択
                selected_option = st.radio(
                    "チケットを選択してください：",
                    options=range(len(ticket_options)),
                    format_func=lambda x: ticket_options[x][0],
                    index=current_selection_idx,
                    key=f"ticket_selection_page_{st.session_state.current_page}"
                )
                
                # 選択されたチケットIDを更新
                if selected_option is not None:
                    new_ticket_id = ticket_options[selected_option][1]
                    if st.session_state.selected_ticket_id != new_ticket_id:
                        st.session_state.selected_ticket_id = new_ticket_id
                
                # 詳細情報用のテーブル表示
                st.markdown("---")
                st.markdown("**詳細情報:**")
                
                # データフレーム表示（情報確認用）
                display_columns = ['ID', '件名', 'ステータス', '優先度', '担当者', '進捗率', '作成日']
                display_df = page_df[display_columns].copy()
                
                # 件名を短縮
                display_df['件名'] = display_df['件名'].apply(
                    lambda x: x[:40] + "..." if len(str(x)) > 40 else str(x)
                )
                
                # 選択されたチケットをハイライト
                if st.session_state.selected_ticket_id:
                    highlight_mask = display_df['ID'] == st.session_state.selected_ticket_id
                    if highlight_mask.any():
                        st.markdown("**選択中のチケット:**")
                        selected_row = display_df[highlight_mask]
                        st.dataframe(selected_row, use_container_width=True)
                        
                        st.markdown("**その他のチケット:**")
                        other_rows = display_df[~highlight_mask]
                        if not other_rows.empty:
                            st.dataframe(other_rows, use_container_width=True)
                    else:
                        st.dataframe(display_df, use_container_width=True)
                else:
                    st.dataframe(display_df, use_container_width=True)
            
            # 選択されたチケットがある場合の詳細表示
            if st.session_state.selected_ticket_id:
                selected_ticket_id = st.session_state.selected_ticket_id
                
                # 選択されたチケットの詳細表示
                st.markdown("---")
                st.markdown(f"### 📋 選択チケット: #{selected_ticket_id}")
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    try:
                        # チケット詳細データを取得
                        ticket_detail = client.get_issue_by_id(selected_ticket_id)
                        
                        # 基本情報を表示
                        info_col1, info_col2 = st.columns(2)
                        
                        with info_col1:
                            st.markdown("**基本情報**")
                            st.write(f"**件名**: {ticket_detail.get('subject', '')}")
                            st.write(f"**トラッカー**: {ticket_detail.get('tracker', {}).get('name', '')}")
                            st.write(f"**ステータス**: {ticket_detail.get('status', {}).get('name', '')}")
                            st.write(f"**優先度**: {ticket_detail.get('priority', {}).get('name', '')}")
                            st.write(f"**進捗率**: {ticket_detail.get('done_ratio', 0)}%")
                        
                        with info_col2:
                            st.markdown("**担当・日程**")
                            st.write(f"**作成者**: {ticket_detail.get('author', {}).get('name', '')}")
                            assigned_to = ticket_detail.get('assigned_to', {}).get('name', '') if ticket_detail.get('assigned_to') else '未設定'
                            st.write(f"**担当者**: {assigned_to}")
                            st.write(f"**開始日**: {ticket_detail.get('start_date', '未設定')}")
                            st.write(f"**期限日**: {ticket_detail.get('due_date', '未設定')}")
                            st.write(f"**実績工数**: {ticket_detail.get('spent_hours', 0)} 時間")
                        
                        # 説明文を表示
                        st.markdown("**説明**")
                        description = ticket_detail.get('description', '説明なし')
                        if description and description.strip():
                            # 長い説明文の場合は折りたたみ表示
                            if len(description) > 300:
                                with st.expander("説明を表示", expanded=False):
                                    st.write(description)
                            else:
                                st.write(description)
                        else:
                            st.write("説明なし")
                        
                        # コメント表示
                        journals = ticket_detail.get('journals', [])
                        comments = [j for j in journals if j.get('notes', '').strip()]
                        
                        if comments:
                            st.markdown("**コメント履歴**")
                            with st.expander(f"コメント ({len(comments)}件)", expanded=False):
                                for comment in comments[-5:]:  # 最新5件まで表示
                                    user_name = comment.get('user', {}).get('name', '不明なユーザー')
                                    created_on = comment.get('created_on', '')
                                    if created_on:
                                        created_on = created_on[:19].replace('T', ' ')
                                    notes = comment.get('notes', '')
                                    
                                    st.markdown(f"**{user_name}** - {created_on}")
                                    st.write(notes)
                                    st.markdown("---")
                        
                    except Exception as e:
                        st.error(f"チケット詳細の取得に失敗しました: {e}")
                    
                    with col2:
                        st.subheader("📄 帳票出力")
                        st.markdown(f"**選択チケット**: #{selected_ticket_id}")
                        
                        if st.button("📋 PowerPoint帳票生成", type="primary", use_container_width=True):
                            try:
                                with st.spinner("PowerPoint帳票を生成中..."):
                                    # 既に取得済みのチケット詳細データを使用
                                    ppt_gen = PowerPointGenerator()
                                    ppt_bytes = ppt_gen.create_issue_report(ticket_detail)
                                    
                                    st.download_button(
                                        label="💾 PowerPointファイルをダウンロード",
                                        data=ppt_bytes,
                                        file_name=f"ticket_{selected_ticket_id}_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                        use_container_width=True
                                    )
                                    
                                    st.success(f"✅ チケット#{selected_ticket_id}の帳票が生成されました！")
                                    
                            except Exception as e:
                                st.error(f"PowerPoint生成エラー: {e}")
            else:
                st.info("表示するチケットがありません。フィルター条件を確認してください。")
        else:
            st.info("チケットデータがありません。")
    
    with tab2:
        st.subheader("CSVエクスポート")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.write(f"現在のフィルター条件でのチケット数: **{len(filtered_df)}件**")
            if not filtered_df.empty:
                st.write("エクスポート対象:")
                st.write(f"- プロジェクト: {selected_project}")
                st.write(f"- ステータス: {selected_status}")
        
        with col2:
            if not filtered_df.empty:
                csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 CSVファイルをダウンロード",
                    data=csv,
                    file_name=f"redmine_tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info("エクスポートするデータがありません。")

def main():
    """メイン関数 - 画面遷移を制御"""
    # セッション状態の初期化
    if 'connected' not in st.session_state:
        st.session_state.connected = False
    
    # 接続状態によって画面を切り替え
    if not st.session_state.connected:
        show_welcome_screen()
    else:
        show_dashboard()

if __name__ == "__main__":
    main()