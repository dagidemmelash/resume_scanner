from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile
import docx2txt
import docx
import os

os.getcwd()


# resume = docx2txt.process('C:\\Users\\fidel\\OneDrive\\Desktop\\Resume\\resume_011 (1).docx')
#
#
# job_description = docx2txt.process('C:\\Users\\fidel\\OneDrive\\Desktop\\Resume\\job_description.docx')



def open_jobdes():
    f = askopenfile(mode='r', filetypes=[('files', ['*.pdf', '*.doc', '*.docx'])])
    if f:
        file_path = os.path.abspath(f.name)
        e1.delete(0, 'end')
        e1.insert(0, str(file_path))
        return file_path


def open_resume():
    f = askopenfile(mode='r', filetypes=[('file', ['*.pdf', '*.doc', '*.docx'])])
    if f:
        file_path = os.path.abspath(f.name)
        e2.delete(0, 'end')
        e2.insert(0, str(file_path))
        return file_path

def match():
    a = str(e1.get())  # job description file path
    b = str(e2.get())  # resume file path
    # call module function here using a and b
    resume = docx2txt.process(a)

    job_description = docx2txt.process(b)

    text = [resume, job_description]

    from sklearn.feature_extraction.text import CountVectorizer

    cv = CountVectorizer()
    count_matrix = cv.fit_transform(text)

    from sklearn.metrics.pairwise import cosine_similarity

    print("/nSimilarity Scores: ")
    print(cosine_similarity(count_matrix))

    matchPercentage = cosine_similarity(count_matrix)[0][1] * 100
    matchPercentage = round(matchPercentage, 2)
    print("Your resume matches about" + str(matchPercentage) + "% of job description.")

    matchdata = str(matchPercentage) + "% Matched"
    result.config(text=matchdata)


def open_jobdes():
    f = askopenfile(mode='r', filetypes=[('files', ['*.pdf', '*.doc', '*.docx'])])
    if f:
        file_path = os.path.abspath(f.name)
        e1.delete(0, 'end')
        e1.insert(0, str(file_path))
        return file_path


def open_resume():
    f = askopenfile(mode='r', filetypes=[('file', ['*.pdf', '*.doc', '*.docx'])])
    if f:
        file_path = os.path.abspath(f.name)
        e2.delete(0, 'end')
        e2.insert(0, str(file_path))
        return file_path


w1 = PanedWindow(height=280, orient=VERTICAL)
w1.pack(fill=BOTH, expand=1)

w2 = PanedWindow(w1, orient=VERTICAL)

w3 = PanedWindow(w2, orient=HORIZONTAL)
w4 = PanedWindow(w2, orient=HORIZONTAL)
w1.add(w2)

job = Label(w3, text="Job: ")
res = Label(w4, text="Resume: ")
f1 = Button(w3, text="Open", command=open_jobdes)
f2 = Button(w4, text="Open", command=open_resume)
e1 = Entry(w3)
e1.insert(0, "enter file path")
e2 = Entry(w4)
e2.insert(0, "enter file path")

w3.add(job)
w3.add(e1)
w3.add(f1)
w4.add(res)
w4.add(e2)
w4.add(f2)

w2.add(w3)
w2.add(w4)
button = Button(w2, text="Match", command=match)
w2.add(button)

labelframe1 = LabelFrame(w1, text="Match result")
labelframe1.pack(pady=20, fill="both", expand=1)
result = Label(labelframe1, font=('Arial', 32), foreground='blue', text="")
result.pack(pady=10)

w1.add(labelframe1)
mainloop()
