from flask import Flask, render_template, jsonify
import threading
import scheduled_crawler  # ✅ 크롤링 코드 import

app = Flask(__name__)

@app.route('/run-crawler', methods=['GET'])
def run_crawler():
    """📌 버튼 클릭 시 크롤링 실행"""

    def run():
        file_name = scheduled_crawler.run_crawler()  # ✅ 크롤링 실행
        if file_name:
            scheduled_crawler.send_email()  # ✅ 크롤링 완료 후 이메일 전송

    thread = threading.Thread(target=run)  # ✅ Flask 서버 블로킹 방지 (백그라운드 실행)
    thread.start()

    return jsonify({"message": "크롤링이 실행되었습니다!"})  # ✅ JSON 응답 반환

if __name__ == '__main__':
    app.run(debug=True)