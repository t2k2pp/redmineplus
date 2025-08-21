import requests
import json

def test_redmine_api():
    base_url = "http://localhost:3000"
    api_key = "d12341dfd7bfa3fabee13d78732989fd913e6bae"
    
    headers = {
        'X-Redmine-API-Key': api_key,
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.get(f"{base_url}/issues.json", headers=headers)
        response.raise_for_status()
        
        data = response.json()
        print("API接続成功!")
        print(f"取得したチケット数: {data.get('total_count', 0)}")
        
        if data.get('issues'):
            print("\n最初のチケット情報:")
            first_issue = data['issues'][0]
            print(json.dumps(first_issue, indent=2, ensure_ascii=False))
        
        return True
    
    except requests.exceptions.RequestException as e:
        print(f"API接続エラー: {e}")
        return False

if __name__ == "__main__":
    test_redmine_api()