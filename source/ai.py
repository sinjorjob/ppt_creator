import openai

def prompt_ai(prompt, api_key):
    openai.api_key = api_key
    completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )
    return completion.choices[0].message.content

#Get the sub topics to a specific topic
def getTopics(topic, language, api_key):
    #msg = f"{topic}についてのPowerPointの資料を作成する必要があります。最も優れたトップカテゴリは何ですか。
    # リスト形式で3から5つのポイントを提案してください。各ポイントについては1～3の単語で表現してください。すべて{language}でお願いします。"
    msg = f"Give me one to three of the most important and generalised topics for a powerpoint presentation about {topic} as bullet points. Only one or two words per point and please in {language}"
    return prompt_ai(msg, api_key).split("\n")

#get all the information to a subtopic 
def getTopicInformations(sub_topic, topic, language, api_key):

    # メイントピック{topic}から{sub_topic}に関する3つまたは4つの最も重要な情報を、8単語以内の短い箇条書きでお願いします。すべて{language}で、わかりやすい言語スタイルで提供してください。
    msg = f"Give me the three or four most important informations about {sub_topic} from the main topic {topic} as  bullet points without numbers and not more than 8 words. Please all in {language} and in a easy understandable language style."
    return prompt_ai(msg, api_key)