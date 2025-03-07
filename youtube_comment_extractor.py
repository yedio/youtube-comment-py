from youtube_comment_downloader import *
import pandas as pd
import os
from datetime import datetime, timezone
import time
import pytz
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLineEdit, QPushButton, QLabel, QRadioButton, 
                            QButtonGroup, QHBoxLayout, QTextEdit)
from PyQt5.QtCore import QThread, pyqtSignal
import sys

class DownloadWorker(QThread):
    progress = pyqtSignal(str)
    
    def __init__(self, url, sort_by):
        super().__init__()
        self.url = url
        self.sort_by = sort_by
        
    def run(self):
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
            self.progress.emit("댓글 다운로드 시작...")
            comments = downloader.get_comments_from_url(self.url, sort_by=self.sort_by)
            for comment in comments:
                timestamp = comment['time_parsed']
                formatted_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d")
                comments_list.append({
                    '닉네임': comment['author'],
                    '날짜': formatted_date,
                    '내용': comment['text']
                })
            
            if not comments_list:
                self.progress.emit("경고: 댓글을 가져오지 못했습니다. 가능한 원인:")
                self.progress.emit("1. 영상의 댓글이 비활성화되어 있습니다.")
                self.progress.emit("2. 영상이 존재하지 않거나 접근할 수 없습니다.")
                self.progress.emit("3. URL이 올바르지 않습니다.")
                return
                
            # DataFrame으로 변환하고 엑셀로 저장
            df = pd.DataFrame(comments_list)
            df.to_excel(output_file, index=False, engine='openpyxl')
            self.progress.emit(f"댓글이 {output_file}에 저장되었습니다.")
            self.progress.emit(f"총 {len(comments_list)}개의 댓글이 저장되었습니다.")
            
        except Exception as e:
            self.progress.emit(f"오류 발생: {str(e)}")

class YoutubeCommentExtractorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.worker = None
        
    def initUI(self):
        self.setWindowTitle('YouTube 댓글 추출기')
        self.setGeometry(100, 100, 600, 400)  # 창 크기 증가
        
        # 중앙 위젯 생성
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 수직 레이아웃 설정
        layout = QVBoxLayout()
        
        # URL 입력 안내 레이블
        url_label = QLabel('YouTube 비디오 URL을 입력하세요:')
        layout.addWidget(url_label)
        
        # URL 입력창
        self.url_input = QLineEdit()
        layout.addWidget(self.url_input)
        
        # 정렬 옵션을 위한 수평 레이아웃
        sort_layout = QHBoxLayout()
        
        # 정렬 옵션 레이블
        sort_label = QLabel('정렬 옵션:')
        sort_layout.addWidget(sort_label)
        
        # 라디오 버튼 그룹 생성
        self.sort_group = QButtonGroup()
        
        # 인기순 라디오 버튼
        self.popular_radio = QRadioButton('인기순')
        self.popular_radio.setChecked(True)  # 기본값으로 선택
        self.sort_group.addButton(self.popular_radio)
        sort_layout.addWidget(self.popular_radio)
        
        # 최신순 라디오 버튼
        self.recent_radio = QRadioButton('최신순')
        self.sort_group.addButton(self.recent_radio)
        sort_layout.addWidget(self.recent_radio)
        
        layout.addLayout(sort_layout)
        
        # 다운로드 버튼
        self.download_btn = QPushButton('댓글 다운로드')
        self.download_btn.clicked.connect(self.start_download)
        layout.addWidget(self.download_btn)
        
        # 로그 출력 영역
        log_label = QLabel('로그:')
        layout.addWidget(log_label)
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)  # 읽기 전용으로 설정
        layout.addWidget(self.log_output)
        
        central_widget.setLayout(layout)
        
    def log(self, message):
        """로그 메시지를 GUI에 추가"""
        self.log_output.append(message)
        # 스크롤을 항상 최하단으로 이동
        self.log_output.verticalScrollBar().setValue(
            self.log_output.verticalScrollBar().maximum()
        )
        
    def start_download(self):
        url = self.url_input.text()
        if url:
            self.log("다운로드 시작...")
            self.download_btn.setEnabled(False)  # 다운로드 중에는 버튼 비활성화
            
            # 정렬 옵션 확인
            sort_by = SORT_BY_POPULAR if self.popular_radio.isChecked() else SORT_BY_RECENT
            
            # 워커 스레드 생성 및 시작
            self.worker = DownloadWorker(url, sort_by)
            self.worker.progress.connect(self.log)
            self.worker.finished.connect(self.download_finished)
            self.worker.start()
        else:
            self.log("URL을 입력해주세요.")
            
    def download_finished(self):
        self.download_btn.setEnabled(True)  # 다운로드 완료 후 버튼 활성화
        self.worker = None

def main():
    app = QApplication(sys.argv)
    ex = YoutubeCommentExtractorGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()