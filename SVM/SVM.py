import nltk
import numpy as np
import pandas as pd

neg = pd.read_excel("neg.xlsx", header=None)
pos = pd.read_excel("pos.xlsx", header=None)
print(neg)
neg['words'] = neg[0].apply(lambda x: nltk.word_tokenize(x))  # separate sentence into words
pos['words'] = pos[0].apply(lambda x: nltk.word_tokenize(x))
x = np.concatenate((pos['words'], neg['words']))
y = np.concatenate((np.ones(len(pos)), np.zeros(len(neg))))

from gensim.models.word2vec import Word2Vec  # word to vec

w2v = Word2Vec(size=300, min_count=5)  # less than 5 time
w2v.build_vocab(x)
w2v.train(x, total_examples=w2v.corpus_count, epochs=w2v.iter)


def total_vec(words):
    vec = np.zeros(300).reshape((1, 300))
    for word in words:
        try:
            vec += w2v.wv[word].reshape((1, 300))
        except KeyError:
            continue
    return vec


train_vec = np.concatenate([total_vec(words) for words in x])

from sklearn.model_selection import cross_val_score
from sklearn.svm import SVC

model = SVC(kernel="rbf", verbose=True)
model.fit(train_vec, y)
scores = cross_val_score(model, x, y, cv=5, scoring='accuracy').mean()
print(scores)


def svm_predict(model):  # 读新评论 进行预测
    df = pd.read_excel("pinglun.xlsx")
    comment_sentiment = []
    for string in df['review_body']:
        words = nltk.word_tokenize(str(string))
        words_vec = total_vec(words)
        result = model.predict(words_vec)
        comment_sentiment.append('good' if int(result[0]) else 'bad')
        if int(result[0]) == 1:
            print(string, '[good]')
        else:
            print(string, '[bad]')


svm_predict(model)
