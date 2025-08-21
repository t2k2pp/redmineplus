from redmine_client import RedmineClient
from ppt_generator import PowerPointGenerator

def test_ppt_generation():
    try:
        # Redmineクライアント初期化
        client = RedmineClient("http://localhost:3000", "d12341dfd7bfa3fabee13d78732989fd913e6bae")
        
        # テスト用チケット取得（ID=368を使用）
        issue_data = client.get_issue_by_id(368)
        
        print("取得したチケット情報:")
        print(f"ID: {issue_data.get('id')}")
        print(f"件名: {issue_data.get('subject')}")
        print(f"トラッカー: {issue_data.get('tracker', {}).get('name')}")
        print(f"ステータス: {issue_data.get('status', {}).get('name')}")
        print(f"コメント数: {len(issue_data.get('journals', []))}")
        
        # PowerPoint生成
        ppt_gen = PowerPointGenerator()
        ppt_bytes = ppt_gen.create_issue_report(issue_data)
        
        # ファイルに保存
        with open("test_ticket_368.pptx", "wb") as f:
            f.write(ppt_bytes)
        
        print(f"\nPowerPointファイルが正常に生成されました: test_ticket_368.pptx")
        print(f"ファイルサイズ: {len(ppt_bytes)} bytes")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    test_ppt_generation()