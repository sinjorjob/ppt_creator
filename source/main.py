import collections.abc
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import powerpoint_helper
import ai
import data_cleaner
import os
import time



def main(topic):
    topic = topic
    author = "siny"
    language = "japanese"
    api_key = "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" 
    #create the presentation from a template to get the correct format (16/9):
    prs = Presentation("sample.pptx")
    # テーマをコピーして新しいプレゼンテーションに適用する

    # TOPスライドの作成(title ,作成者)
    powerpoint_helper.add_top_slide(prs, topic, f"By {author}",48, "Meiryo" ,PP_ALIGN.CENTER)

    print("DONE")

    #目次スライドの作成
    """
    1. コミュニケーション
    2. 生産性向上
    3. アウトソーシング
    4. 統合コミュニケーション
    5. リアルタイムコラボレーション
    6. クラウドストレージ
    7. セキュリティ
    8. カスタマーサポート
    """
    sub_topics = ai.getTopics(topic, language, api_key)
    print(sub_topics)

    #フォーマットの修正、文字列型に変換
    clean_data = data_cleaner.clean_bullet_points("\n".join(sub_topics)) 

    print(clean_data)

    #table of content:
    powerpoint_helper.add_text_slide2(prs, "目次", clean_data ,48, "Meiryo" ,PP_ALIGN.LEFT)


    # サブトピックをループして、各トピックについてのスライドを作成

    count = 0
    for sub_topic in sub_topics:
        count = count + 1
        print(f"Slide {count}/{len(sub_topics)}")
        print("wait 20 second...")
        time.sleep(20)
        data = ai.getTopicInformations(sub_topic, topic, language, api_key)
        print("サブコンテンツ＝", type(data), data)
        clean_data = data_cleaner.clean_bullet_points(data)
        print("サブコンテンツ2＝", type(clean_data), clean_data)
        powerpoint_helper.add_text_slide2(prs, sub_topic, clean_data,48, "Meiryo" ,PP_ALIGN.LEFT)
    prs.save(f"{topic}.pptx")




if __name__ == '__main__':
    topic = input("資料作成したいトピックを入力してください:\n")
    main(topic)
