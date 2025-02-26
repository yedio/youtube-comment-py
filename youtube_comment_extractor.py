from youtube_comment_downloader import *
import pandas as pd
import os
from datetime import datetime, timezone
import time
import pytz

def download_comments(url):
    # 결과물 저장할 폴더 생성
    output_dir = "youtube_comments"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 파일명에 초까지 포함
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_dir, f"youtube_comments_{current_time}.xlsx")
    
    # 댓글 다운로더 인스턴스 생성
    downloader = YoutubeCommentDownloader()
    
    # 댓글 수집
    comments_list = []
    try:
        # 디버깅을 위한 출력 추가
        print("댓글 다운로드 시작...")
        comments = downloader.get_comments_from_url(url, sort_by=SORT_BY_POPULAR)
        for comment in comments:
            timestamp = comment['time_parsed']
            formatted_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d")
            comments_list.append({
                '닉네임': comment['author'],
                '날짜': formatted_date,
                '내용': comment['text']
            })
        
        if not comments_list:
            print("경고: 댓글을 가져오지 못했습니다. 가능한 원인:")
            print("1. 영상의 댓글이 비활성화되어 있습니다.")
            print("2. 영상이 존재하지 않거나 접근할 수 없습니다.")
            print("3. URL이 올바르지 않습니다.")
            return
            
        # DataFrame으로 변환하고 엑셀로 저장
        df = pd.DataFrame(comments_list)
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"댓글이 {output_file}에 저장되었습니다.")
        print(f"총 {len(comments_list)}개의 댓글이 저장되었습니다.")
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")

def main():
    url = input("YouTube 비디오 URL을 입력하세요: ")
    print("댓글을 다운로드하는 중...")
    download_comments(url)

if __name__ == "__main__":
    main()