import docx
import telebot
import re
import os
import codecs

from nltk import PorterStemmer
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize, word_tokenize

bot_token = "token"
bot = telebot.TeleBot(bot_token)


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Привет! Выгрузи документ, и я скажу тебе, о чем он. (Не больше 20МБ)')


@bot.message_handler(commands=['help'])
def start_message(message):
    bot.send_message(message.chat.id,
                     "Привет! Я бот, который анализирует текстовые документы и выводит аннотацию к ним, так что просто загрузи в этот диалог документ, и я проанализирую его! (Не больше 20МБ)")


def error(message):
    bot.send_message(message.chat.id,
                     "Я бот, который позволяет проанализировать текстовый документ и вывести аннотацию к нему, так что просто загрузи в этот диалог документ, и я проанализирую его! (Не больше 20МБ)")


@bot.message_handler(content_types=["audio", "photo", "text"])
def main(message):
    error(message)


@bot.message_handler(content_types=['document'])
def documentProcessig(message):

    chat_id = message.chat.id
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        src = u'C:\\Users\\Kangarooo444\\PycharmProjects\\DiplomTextSum\\downloads\\' + message.document.file_name;
        with open(src, 'wb') as new_file:
            new_file.write(downloaded_file)
    except:
        bot.send_message(chat_id, 'Произошла ошибка! Вес файла превышает 20 МБ.')
        return
    try:
        fileObj = codecs.open(src, "r", "utf_8_sig")
        Res = TextSum(fileObj.read())  # Открытие как .txt
        fileObj.close()
    except:
        try:
            Res = TextSum(getText(src))  # Открытие как .doc
        except:
            bot.send_message(chat_id, "Не удалось корректно открыть файл")
            return
        finalMsg = "<b><I>Резюме по документу:</I></b> \n\n" + Res

        # Проверка длины итогового текста и его отправка
        if len(finalMsg) > 4096:
            for x in range(0, len(finalMsg), 4096):
                bot.send_message(message.chat.id, finalMsg[x:x + 4096], parse_mode='HTML')
            else:
                bot.send_message(message.chat.id, finalMsg, parse_mode='HTML')
        os.remove(src)


# Получение текста из .docx
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


# Суммаризация текста

def TextSum(text_string):
    text_string = re.sub(r'\[[0-9]*\]', ' ', text_string)
    text_string = re.sub(r's+', ' ', text_string)

    formatted_input_text = re.sub('[^а-яА-Я]', ' ', text_string)
    formatted_input_text = re.sub(r's+', ' ', formatted_input_text)

    stopWords = set(stopwords.words("russian"))
    ps = PorterStemmer()

    # Удаление стоп-слов

    freqTable = dict()
    for word in word_tokenize(formatted_input_text, language="russian"):
        word = ps.stem(word)
        if word in stopWords:
            continue

        if word in freqTable:
            freqTable[word] += 1
        else:
            freqTable[word] = 1

    sentences = sent_tokenize(text_string)  # Разделение всего текста на предложения
    sentenceValue = dict()

    # Расчет ценности предложений

    for sentence in sentences:
        word_count_in_sentence = (len(word_tokenize(sentence)))
        for wordValue in freqTable:
            if wordValue in sentence.lower():
                if sentence[:10] in sentenceValue:
                    sentenceValue[sentence[:10]] += freqTable[wordValue]
                else:
                    sentenceValue[sentence[:10]] = freqTable[wordValue]

    sentenceValue[sentence[:10]] = sentenceValue[sentence[:10]] // word_count_in_sentence

    sumValues = 0
    for entry in sentenceValue:
        sumValues += sentenceValue[entry]

    average = int(sumValues / len(sentenceValue))

    sentence_count = 0
    summary = ''

    # Удаление предложений с наименьшей ценностью

    for sentence in sentences:
        if sentence[:10] in sentenceValue and sentenceValue[sentence[:10]] > (average):
            summary += " " + sentence
            sentence_count += 1
            
    return summary


bot.polling(none_stop=True, interval=0)
