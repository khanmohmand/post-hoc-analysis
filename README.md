# Post Hoc Analysis System

Streamlit app jo MCQ response sheets se Post Hoc (item) analysis nikalti hai:
Calculations, Report, All 3 / All 4 Column Questions aur NFD (Non Functional Distractor).

## Excel file ki shakal
- **Responses sheet:** pehli row mein `Q1, Q2, Q3 ...`, neeche har row ek student (A/B/C/D/E).
- **Keys sheet:** ya to `QNO` aur `KEY` columns, ya pehli row mein `Q1, Q2 ...` aur doosri row mein sahi jawab.
- Sheets ki tarteeb koi bhi ho sakti hai. Questions kitne bhi ho sakte hain.

## Apne computer par chalana
```
python -m pip install -r requirements.txt
python -m streamlit run post_hoc_app.py
```

## Internet par link banana (Streamlit Community Cloud)
1. GitHub par naya repository banayein (naam jaise `post-hoc-analysis`).
2. Is folder ki teeno files upload karein: `post_hoc_app.py`, `requirements.txt`, `README.md`.
3. https://share.streamlit.io par GitHub se sign in karein, **Create app** dabayein.
4. Repository, branch (`main`) aur main file `post_hoc_app.py` chunein, phir **Deploy**.
5. Kuch minute baad aap ko ek permanent link milta hai jo colleagues ke sath share kar sakte hain.

Code mein tabdeeli GitHub par file edit karke save karne se link wala app khud update ho jata hai.

## Dhyan rakhein
Exam ka data confidential hota hai. Public link par upload ki hui files server par process hoti hain.
App ki Share settings mein link ko sirf chune hue logon tak mehdood karein, ya university ki IT policy ke
mutabiq sirf office ke computer par chalayein.
