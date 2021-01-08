import matplotlib.pyplot as plt
import textract
from collections import Counter

filename = "thesis.docx"
totalWords = 5
commonWords = ["i", "me", "my", "myself", "we", "our", "ours", "ourselves",
        "you", "your", "yours", "yourself", "yourselves", "he", "him", "his",
        "himself", "she", "her", "hers", "herself", "it", "its", "itself",
        "they", "them", "their", "theirs", "themselves", "what", "which", "who",
        "whom", "this", "that", "these", "those", "am", "is", "are", "was",
        "were", "be", "been", "being", "have", "has", "had", "having", "do",
        "does", "did", "doing", "a", "an", "the", "and", "but", "if", "or",
        "because", "as", "until", "while", "of", "at", "by", "for", "with",
        "about", "against", "between", "into", "through", "during", "before",
        "after", "above", "below", "to", "from", "up", "down", "in", "out",
        "on", "off", "over", "under", "again", "further", "then", "once",
        "here", "there", "when", "where", "why", "how", "all", "any", "both",
        "each", "few", "more", "most", "other", "some", "such", "no", "nor",
        "not", "only", "own", "same", "so", "than", "too", "very", "s", "t",
        "can", "will", "just", "don", "should", "now"]

def cleanWord(word : str) -> str:
    word = word.decode('utf-8')
    if word.isalnum():
        return word.lower()

### Get text from Word document 
def getText(filename : str) -> list:
    text = textract.process(filename)
    text = text.split()
    text = list(map(cleanWord, text))
    return [word for word in text if word != None]

### Filter common words 
thesisUncommon = [word for word in getText(filename) if word not in commonWords]


wordFrequencyPairs = dict(Counter(thesisUncommon)).items()
sortedWordFrequency = sorted(wordFrequencyPairs,  key=lambda item: item[1], reverse=True)

labels = [items[0] for items in sortedWordFrequency[0:totalWords]]
frequency = [items[1] for items in sortedWordFrequency[0:totalWords]]

# Creating a custom autopct for the pie chart
def make_autopct(frequency):
    def custom_autopct(pct):
        total = sum(frequency)
        val = int(round(pct*total/100.0))
        return str(val)
    return custom_autopct

# Plt to create the plot
fig, (ax1, ax2) = plt.subplots(1, 2)

# Style points
centerCircle = plt.Circle((0,0),0.70,fc='white')
centerCircle2 = plt.Circle((0,0),0.70,fc='white')
ax1.add_artist(centerCircle)
ax2.add_artist(centerCircle2)

# Creating ax1
ax1.set_title("Most Common Words")

ax1.pie(
        frequency,
        labels=labels,
        autopct=make_autopct(frequency),
        )


wordsAsSet = set(thesisUncommon)
wordSizes = [(word, len(word)) for word in wordsAsSet]
longestWordsAndSize = sorted(wordSizes, key=lambda item: item[1], reverse=True)
longestWords = [item[0] for item in longestWordsAndSize[0:totalWords]]
sizes = [item[1] for item in longestWordsAndSize[0:totalWords]]

# Creating ax2
ax2.set_title("Longest words used")
ax2.pie(
        sizes,
        labels=longestWords,
        autopct=make_autopct(sizes),
        )

plt.show()

