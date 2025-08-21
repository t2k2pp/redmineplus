import requests
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional

class RedmineClient:
    def __init__(self, base_url: str, api_key: str):
        self.base_url = base_url.rstrip('/')
        self.api_key = api_key
        self.headers = {
            'X-Redmine-API-Key': api_key,
            'Content-Type': 'application/json'
        }
    
    def get_issues(self, limit: int = 100, offset: int = 0, **kwargs) -> Dict:
        params = {
            'limit': limit,
            'offset': offset,
            **kwargs
        }
        
        try:
            response = requests.get(
                f"{self.base_url}/issues.json",
                headers=self.headers,
                params=params
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            raise Exception(f"チケット取得エラー: {e}")
    
    def get_all_issues(self) -> List[Dict]:
        all_issues = []
        offset = 0
        limit = 100
        
        while True:
            data = self.get_issues(limit=limit, offset=offset)
            issues = data.get('issues', [])
            
            if not issues:
                break
                
            all_issues.extend(issues)
            
            if len(issues) < limit:
                break
                
            offset += limit
        
        return all_issues
    
    def get_issue_by_id(self, issue_id: int) -> Dict:
        try:
            params = {
                'include': 'journals'
            }
            response = requests.get(
                f"{self.base_url}/issues/{issue_id}.json",
                headers=self.headers,
                params=params
            )
            response.raise_for_status()
            return response.json()['issue']
        except requests.exceptions.RequestException as e:
            raise Exception(f"チケット詳細取得エラー: {e}")
    
    def get_projects(self) -> List[Dict]:
        try:
            response = requests.get(
                f"{self.base_url}/projects.json",
                headers=self.headers
            )
            response.raise_for_status()
            return response.json().get('projects', [])
        except requests.exceptions.RequestException as e:
            raise Exception(f"プロジェクト取得エラー: {e}")
    
    def get_users(self) -> List[Dict]:
        try:
            response = requests.get(
                f"{self.base_url}/users.json",
                headers=self.headers
            )
            response.raise_for_status()
            return response.json().get('users', [])
        except requests.exceptions.RequestException as e:
            raise Exception(f"ユーザー取得エラー: {e}")
    
    def issues_to_dataframe(self, issues: List[Dict]) -> pd.DataFrame:
        data = []
        
        for issue in issues:
            row = {
                'ID': issue.get('id'),
                'プロジェクト': issue.get('project', {}).get('name', ''),
                'トラッカー': issue.get('tracker', {}).get('name', ''),
                'ステータス': issue.get('status', {}).get('name', ''),
                '優先度': issue.get('priority', {}).get('name', ''),
                '件名': issue.get('subject', ''),
                '説明': issue.get('description', ''),
                '作成者': issue.get('author', {}).get('name', ''),
                '担当者': issue.get('assigned_to', {}).get('name', '') if issue.get('assigned_to') else '',
                '開始日': issue.get('start_date', ''),
                '期限日': issue.get('due_date', ''),
                '進捗率': issue.get('done_ratio', 0),
                '予定工数': issue.get('estimated_hours', 0) or 0,
                '実績工数': issue.get('spent_hours', 0) or 0,
                '作成日': issue.get('created_on', ''),
                '更新日': issue.get('updated_on', ''),
                '終了日': issue.get('closed_on', ''),
                'プライベート': issue.get('is_private', False)
            }
            data.append(row)
        
        df = pd.DataFrame(data)
        
        for date_col in ['開始日', '期限日', '作成日', '更新日', '終了日']:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        return df