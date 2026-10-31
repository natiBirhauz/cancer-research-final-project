import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict

# 1. קריאת קובץ האקסל
try:
    df = pd.read_excel('Book1.xlsx')
except FileNotFoundError:
    print("שגיאה: הקובץ 'Book1.xlsx' לא נמצא.")
    exit()

# יצירת מילון חכם שישמור לכל תווית גם את הכמות וגם רשימה של ROIs
# שימוש ב-set מבטיח שאם תווית מופיעה פעמיים באותו ROI, שמו של ה-ROI ייכתב רק פעם אחת
es_data = defaultdict(lambda: {'count': 0, 'rois': set()})

# 2. חילוץ התוויות וה-ROI שורה אחר שורה
for index, row in df.iterrows():
    roi = row['Tail_ROI']
    labels_str = str(row['PartB_ES_Labels'])
    
    # דילוג על שורות ריקות
    if pd.isna(row['PartB_ES_Labels']) or labels_str.strip() == 'nan':
        continue
        
    # פיצול התוויות
    for label in labels_str.split(';'):
        clean_label = label.strip()
        if clean_label:
            es_data[clean_label]['count'] += 1
            es_data[clean_label]['rois'].add(roi)

# 3. הכנת הנתונים לטבלה ומיון
records = []
for es_label, data in es_data.items():
    # הופכים את ה-set של ה-ROIs למחרוזת מסודרת שמופרדת בפסיקים
    rois_joined = ', '.join(sorted([str(r) for r in data['rois'] if str(r) != 'nan']))
    
    records.append({
        'ES_Label': es_label,
        'Frequency': data['count'],
        'Tail_ROIs': rois_joined # העמודה החדשה!
    })

# יצירת DataFrame ומיון מהנפוץ ביותר לנדיר ביותר
ranking_df = pd.DataFrame(records)
ranking_df = ranking_df.sort_values(by='Frequency', ascending=False).reset_index(drop=True)

# 4. ייצוא לטבלה
output_filename = 'ES_Labels_Ranking_with_ROI.xlsx'
ranking_df.to_excel(output_filename, index=False)
print(f"הטבלה יוצאה בהצלחה לקובץ: {output_filename}")

# 5. יצירת גרף של 10 המובילים
top_10 = ranking_df.head(10)

plt.figure(figsize=(10, 6))
# שימוש בגרסה המעודכנת של seaborn ללא שגיאות עתידיות (FutureWarning)
sns.barplot(x='Frequency', y='ES_Label', data=top_10, hue='ES_Label', palette="viridis", legend=False)

# עיצוב הגרף
plt.title('Top 10 Most Common ES Labels', fontsize=16)
plt.xlabel('Frequency (תדירות)', fontsize=12)
plt.ylabel('ES Label (תווית)', fontsize=12)
plt.tight_layout()

# שמירת הגרף כתמונה והצגתו למשתמש
plt.savefig('top_es_labels.png')
print("הגרף נשמר בהצלחה כקובץ 'top_es_labels.png'")
plt.show()