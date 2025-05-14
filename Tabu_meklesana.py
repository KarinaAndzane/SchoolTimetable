import random
import pandas as pd
import re
import time
from openpyxl import load_workbook
import copy

#start = time.perf_counter()

# Metode, lai iegūtu nejaušu (ne-tukšu) elementu no risinājuma (solution)
def get_random_element_from_solution(solution, clases):
    while True:
        day = random.choice(list(solution.keys()))
        lesson = random.randint(0, 7)  
        class_idx = random.randint(0, len(clases) - 1)

        element = solution[day][lesson][class_idx]
        if element is not None:
            return day, clases[class_idx], lesson + 1, element  

# Sākumdatu apstrāde: skolotāju piešķiršāna klasēm
def teachers_assigment():
    # Lai glābt skolotāju brīvu stundu skaitu
    teacher_hours = {teacher: 23 for teacher in rooms.keys()}
    pinned_teachers = {} 

    for c in clases:
        pinned_teachers[c] = {}
        for s in programs[c[0]]:
            # Jo Klases stundai jau ir piešķirti noteikti skolotāji
            if s == 'Klases stunda': 
                continue

            teacher = ""
            s_count = programs[c[0]][s]
            max = 0
            t_list = []

            for t in teachers[s]:
                if teacher_hours[t] >= s_count:
                    t_list.append(t)
                    
            if len(t_list) == 0:
                print ("Nav pieejamu skolotaju priekšmetam ",s)   

            else:
                for t2 in t_list:
                    if teacher_hours[t2] > max:
                        max = teacher_hours[t2]
                        teacher = t2
            
            teacher_hours[teacher] -= s_count
            pinned_teachers[c][s] = teacher

    return pinned_teachers

# Sākumrisinājuma ģenerēšana
def initial_solution():
   # Skolotāju un telpu pieejamība
    teacher_busy = {
        teacher: {d: [False for _ in range(8)] for d in range(1, 6)}
        for teacher in all_teachers
    }
    room_busy = {
        room: {d: [False for _ in range(8)] for d in range(1, 6)}
        for room in all_rooms
    }
    
    # Risinājuma aizpildīšana sākot no pirmdienas pirmas stundas
    for d in range(1, 6):
        for slot in range(1, 9):
            for c in clases:
                # Maksimals stundu skaits 7.klasem ir 7
                if c.startswith("7") and slot > 7:
                    continue  
                
                s_list = programs_by_class[c]
                # Varbūtības maiņa, lai atvieglot risinājumu veidošānu
                # Jo matemātikas stundas ir visvairāk, bet sporta stundas notiek vienā telpā
                if 'Matemātika'  in s_list:
                    subject_weights = list(s_list.keys()) + ['Matemātika'] *2
                elif 'Sports' in s_list:
                    subject_weights = list(s_list.keys()) + ['Sports'] * 2
                else:
                    subject_weights = list(s_list.keys())

                # Ja visi priekšmeti ir pieškirti, beigt darbu
                if len(s_list) == 0:
                    continue

                attempts = 0
                while attempts < 50:
                    # Nejauši izvēlēts priekšmets
                    r_subject = random.choice(subject_weights)
                    choice = []

                    if r_subject in ["Vācu valoda", "Franču valoda"]:
                        # Divas stundas vienā laikā, izņēmums svešvalodas
                        t1 = pinned_teachers[c]["Vācu valoda"]
                        t2 = pinned_teachers[c]["Franču valoda"]
                        r1 = rooms[t1]
                        r2 = rooms[t2]

                        if (
                            not  teacher_busy[t1][d][slot-1] and
                            not  teacher_busy[t2][d][slot-1] and
                            not  room_busy[r1][d][slot-1] and
                            not  room_busy[r2][d][slot-1]):

                            choice.append(("Vācu valoda", t1, r1))
                            choice.append(("Franču valoda", t2, r2))

                            teacher_busy[t1][d][slot-1] = True
                            teacher_busy[t2][d][slot-1] = True
                            room_busy[r1][d][slot-1] = True
                            room_busy[r2][d][slot-1] = True

                            for lang in ["Vācu valoda", "Franču valoda"]:
                                if s_list[lang] == 1:
                                    del s_list[lang]
                                else:
                                    s_list[lang] -= 1

                    else:
                        if r_subject == "Klases stunda":
                            choice = r_subject, class_teachers[c], rooms[class_teachers[c]]
                        else:
                            t = pinned_teachers[c][r_subject]
                            choice = r_subject, t, rooms[t]

                        
                        # Ja priekšmets tik piešķirts laikam, tad jāizmaina skolotāju un telpu pieejamību
                        if not teacher_busy[choice[1]][d][slot-1] and not room_busy[choice[2]][d][slot-1]:
                            teacher_busy[choice[1]][d][slot-1]= True
                            room_busy[choice[2]][d][slot-1] = True

                            if s_list[r_subject] == 1:
                                del s_list[r_subject]
                            else:
                                s_list[r_subject] -= 1
                        else:
                            choice=[]
                            
                    if choice:
                        c_index = clases.index(c)
                        solution[d][slot-1][c_index] = choice
                        choice = []
                        break
                    else:
                        attempts += 1 

    return teacher_busy, room_busy

# Risinājuma novērtēšana
def evaluate_schedule(solution, clases):
    # 10 vērtības, ja stunda sarakstā ir "logs"
    gap_penalty = 10
    # 3 vērtības, ja vienā dienā atkārtojas viens priekšmets
    repeated_subject_penalty = 3
    # 4 vērtības, ja stundas nedēļā nav vienmērīgi sadalītas pa dienām
    daily_lesson_penalty = 4
    # 6 vērtības, ja skolotāja obligāts nosacījums nav izpildīts (1.privileģijas tips)
    busy_teacher_penalty = 6
    # 1 vērtība, ja skolotāja vēlāms nosacījums nav izpildīts (2.privileģijas tips)
    second_preference_p = 1

    total_penalty = 0
    conflicts_list = []
    repeat_conflicts = []  # (day, class_name, lesson, subject)
    day_names = {1: "Pirmdiena", 2: "Otrdiena", 3: "Trešdiena", 4: "Ceturtdiena", 5: "Piektdiena"}

    num_days = len(solution)
    num_lessons = len(solution[1])  

    for class_idx, class_name in enumerate(clases):
        for day in range(1, num_days + 1):
            day_schedule = [solution[day][lesson][class_idx] for lesson in range(num_lessons)]

            # Obligāta stunda pirmajā stundā
            if day_schedule[0] is None:
                total_penalty += gap_penalty

            # Pārbaude: uz brīviem laika logiem
            first = next((i for i, v in enumerate(day_schedule) if v is not None), None)
            last = next((i for i, v in reversed(list(enumerate(day_schedule))) if v is not None), None)

            if first is not None and last is not None:
                for i in range(first, last + 1):
                    if day_schedule[i] is None:
                        total_penalty += gap_penalty

            # Pārbuade: Vienā dienā atkārtojas viens priekšmets
            subject_to_lessons = {}
            for lesson, entry in enumerate(day_schedule):
                if entry is not None:
                    subject = entry[0] if isinstance(entry, list) else entry
                    if subject not in subject_to_lessons:
                        subject_to_lessons[subject] = []
                    subject_to_lessons[subject].append((lesson, subject))

            for subject, lesson_tuples in subject_to_lessons.items():
                if len(lesson_tuples) > 1:
                    penalty = (len(lesson_tuples) - 1) * repeated_subject_penalty
                    total_penalty += penalty
                    for lesson, subj in lesson_tuples:
                        repeat_conflicts.append((day, class_name, lesson + 1, subj))

            # Pārbaude: stundas nedēļā nav vienmērīgi sadalītas pa dienām
            actual_count = sum(1 for x in day_schedule if x is not None)

            if class_name.startswith("7"):
                min_required = 5
            elif class_name.startswith("8") or class_name.startswith("9"):
                min_required = 6
            else:
                min_required = 5  

            if actual_count < min_required:
                penalty = (min_required - actual_count) * daily_lesson_penalty
                total_penalty += penalty

            # Skolotāju pieejamības pārbaude
            for lesson_index, entry in enumerate(day_schedule):
                if entry is None:
                    continue

                entries = entry if isinstance(entry, list) else [entry]
                for e in entries:
                    if not isinstance(e, (tuple, list)) or len(e) < 3:
                        continue

                    _, teacher, _ = e
                    day_str = day_names[day]
                    busy_slots = dict1.get(teacher, {}).get(day_str, [])
                    secondpreference = dict2.get(teacher, {}).get(day_str, [])

                    if (lesson_index + 1) in busy_slots:
                        total_penalty += busy_teacher_penalty
                        repeat_conflicts.append((day, class_name, lesson_index + 1, solution[day][lesson_index][class_idx]))
                        
                    if (lesson_index + 1) in secondpreference:
                        total_penalty += second_preference_p
                        repeat_conflicts.append((day, class_name, lesson_index + 1, solution[day][lesson_index][class_idx]))
                    
    conflicts_list = repeat_conflicts
    return total_penalty, conflicts_list

# Risinājuma pārbaude uz stingriem ierobežojumiem
def check_hard_constraints(solution):
    teacher_slots = set()
    class_slots = set()
    num_classes = len(solution[1][0])  

    for day, lessons in solution.items():
        for lesson_index, lesson in enumerate(lessons): 
            for class_idx, subject_entry in enumerate(lesson):
                if subject_entry is None:
                    continue

                subjects = subject_entry if isinstance(subject_entry, list) else [subject_entry]

                for subject in subjects:
                    if not isinstance(subject, (tuple, list)) or len(subject) < 3:
                        continue

                    _, teacher, class_name = subject

                    # Stingrs ierobežojums: skolotājs nevar vienlaikus vadīt divas mācību stundas
                    teacher_key = (day, lesson_index, teacher)
                    if teacher_key in teacher_slots:
                        return False
                    else:
                        teacher_slots.add(teacher_key)

                   # Stingrs ierobežojums: klase nevar būt divās mācību stundās vienlaikus
                    class_key = (day, lesson_index, class_name)
                    if class_key in class_slots:
                        return False
                    class_slots.add(class_key)                

        # Pārbaude: pirmajai stundai jābūt obligāti aizpildītai
    for class_idx in range(num_classes):
        for day in solution:
            entry = solution[day][0][class_idx]
            if entry is None:
                return False
            
            entry = solution[day][7][class_idx] 
            
            if class_idx <= 2 and entry is not None: # 1, ja divas 7.klases
                return False


    return True

# Labāka kaimiņa atrašānas funckija
def swap_and_evaluate(solution, class_name, fixed_day, fixed_lesson, clases):
    best_solution = None
    best_subject= None
    best_penalty = float('inf')
    
    class_idx = clases.index(class_name)
    num_days = len(solution)
    num_lessons = len(solution[1])

    # Meklējot labāko kaimiņu, tika ņemta vērā padotā diena, laiks un priekšmets
    fixed_subject = solution[fixed_day][fixed_lesson-1][class_idx]
    current_p, _ = evaluate_schedule(solution, clases)

    # Visi iespejāmi varianti
    for day in range(1, num_days + 1):
        for s in range(num_lessons):
            if day == fixed_day and s == fixed_lesson-1:
                continue 

            new_solution = copy.deepcopy(solution)
            other_subject = new_solution[day][s][class_idx]

            # Maiņa
            new_solution[fixed_day][fixed_lesson-1][class_idx] = other_subject
            new_solution[day][s][class_idx] = fixed_subject

            # Vai pārkapj stingrus ierobežojumus?
            if not check_hard_constraints(new_solution):

                continue  

            # Novertējums
            penalty, _ = evaluate_schedule(new_solution, clases)

            # Ja jauns risinājums labāks, atstājam viņu
            if penalty < best_penalty:
                best_penalty = penalty
                best_solution = new_solution
                best_subject = day,s,class_idx, (other_subject)

    # Ja atbilstošs risinājums nav atrasts, funkcija atgriež sākotnējo risinājumu
    if best_solution is None:
        best_solution = solution
    
    elif best_penalty > current_p:
        best_solution = solution

    return best_solution, best_penalty, best_subject


start_time = time.time()
# Datu lasīšana no Excel faila (klases, skolotāji, telpas, klases skolotāji)
input  = input("Ievadi ievaddatu MS Excel faila nosaukumu: ")
excel_p = input + ".xlsx"
dfKS = pd.read_excel(excel_p, sheet_name="Klases sk.")
dfS = pd.read_excel(excel_p, sheet_name="Skolotaji")
dfP = pd.read_excel(excel_p, sheet_name="Programma")

# KLašu saraksts
class_teachers = dict(zip(dfKS["Klase"], dfKS["Skolotājs"]))
clases = list(class_teachers.keys())

#  solution[diena][laiks][klases_indeks] — vārdnīca ar 3 laukiem: prieksmets, skolotajs, kabinets.
solution = {
    1: [  # Pirmdiena
        [None for _ in clases] for _ in range(1, 9)
    ],
    2: [  # Otrdiena
        [None for _ in clases] for _ in range(1, 9)
    ],
    3: [  # Trešdiena
        [None for _ in clases] for _ in range(1, 9)
    ],
    4: [  # Ceturtdiena
        [None for _ in clases] for _ in range(1, 9)
    ],
    5: [  # Piektdiena
        [None for _ in clases] for _ in range(1, 9)
    ]
}

day_map = {
    'Pirmdiena': 1,
    'Otrdiena': 2,
    'Trešdiena': 3,
    'Ceturtdiena': 4,
    'Piektdiena': 5
}
# Programma
dfP.columns = ['Subject', '7', '8', '9']

# Klase un programma, iszlēdzot priekšmetus, kur ir 0 stundas
programs = {
    "7": {subject: count for subject, count in zip(dfP['Subject'], dfP['7'].astype(int)) if count > 0},
    "8": {subject: count for subject, count in zip(dfP['Subject'], dfP['8'].astype(int)) if count > 0},
    "9": {subject: count for subject, count in zip(dfP['Subject'], dfP['9'].astype(int)) if count > 0}
}

# Katras klases programmas saraksts
programs_by_class = {}
for klase in clases:
    level = klase[0]  
    programs_by_class[klase] = programs[level].copy()
    
# Mācību priekšmets un skolotāji
teachers = {}

# Skolotāju datnes aizpildīšana
for _, row in dfS.iterrows():
    for column in dfS.columns[3:]:
        subject = row[column]
        if pd.notna(subject): 
            if subject not in teachers:
                teachers[subject] = []
            teachers[subject].append(row['Skolotājs'])

teachers["Klases stunda"] = list(class_teachers.values())

# Skolotājs un telpa
rooms = dict(zip(dfS["Skolotājs"], dfS["Telpa"]))

# Visi skolotāji un telpas
all_teachers = dfS["Skolotājs"].dropna().unique().tolist()
all_rooms = dfS["Telpa"].dropna().unique().tolist()


xlsx = pd.ExcelFile(excel_p)
teachers_df = pd.read_excel(excel_p, sheet_name="Skolotaji")
# Vārdnīca, kur atslēga ir skolotāja numurs, bet vērtība ir V. Uzvārds
teacher_names = dict(zip(teachers_df['Nr.'], teachers_df['Skolotājs']))

# Prioritātes grupas
dict1 = {}
dict2 = {}

# Aiz "Klases sk." sākas skolotāju lapas
start_processing = False
for sheet in xlsx.sheet_names:
    if sheet == "Klases sk.":
        start_processing = True
        continue
    if not start_processing:
        continue

    # Excel lapas nosakukma pārbaude (Skolotāja_numurs(prioritāte))
    match = re.match(r"^(\d+)\((1|2)\)$", sheet)
    if not match:
        print(f"Ignorē lapu: {sheet}")
        continue

    teacher_number = int(match.group(1))
    priority = match.group(2)

    # Meklējam skolotāja vārdu
    if teacher_number in teacher_names:
        full_name = teacher_names[teacher_number]
    else:
        print(f"Skolotājs ar numuru {teacher_number} nav atrast")
        continue

    valid_days = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturtdiena", "Piektdiena"]

    df = pd.read_excel(excel_p, sheet_name=sheet)
    df = df.drop(df.columns[0], axis=1)
    df = df.loc[:, df.columns.isin(valid_days)]
    df = df.head(8)

    # Atrodam stundas, kur ir "x" (nevēlamie laiki)
    schedule = {}
    for day in valid_days:
        if day in df.columns:
            hours = df[df[day] == 'x'].index + 1
            schedule[day] = list(hours)

        # Saglābajam datus atbilstošā grupā 
        if priority == '1':
            dict1[full_name] = schedule
        elif priority == '2':
            dict2[full_name] = schedule

    
#  SKOLOTĀJU PIEŠĶIRŠANA KLASĒM
pinned_teachers = teachers_assigment()

#  SĀKONTĒJA RISINĀJUMA ĢENERĒŠANA
initial_solution()

#print(programs_by_class) # pārbaude: jābūt tukš

if  not all(not d for d in programs_by_class.values()):
    pinned_teachers = teachers_assigment()
    teacher_busy, room_busy = initial_solution()
    print(programs_by_class) #pārbaude (jābūt tukšš)

# Maksimāls iteraciju skaits, kad tiek sasniegts, pabeidz darbību
MaxIter = 2000
itera = 0

conflicts_list = []

# Sākotnēja risinājuma novertēšana
conflicts, conflicts_list = evaluate_schedule(solution, clases)
tabu_list = []
current_penalty = 0 
m_counter = 0

while itera != MaxIter:
    itera +=1
    # Ja risinājuma novertējums nemainās 15 iteracijas, tad paņemt elementu no konfiktu sarakstā
    if m_counter < 15:
        element = get_random_element_from_solution(solution, clases)
    else:
        if len(conflicts_list)>=1:
            element = random.choice(list(conflicts_list))
        m_counter = 0
        
    c = element[1]
    s = element[3][0]
    class_idx = clases.index(c)

    # Meklējam labāko kaimiņu
    best_solution, best_penalty, best_subject = swap_and_evaluate(solution, c, element[0], element[2], clases)

    if best_subject is not None:
        tabu_element = ((c, (element[0], element[2]),( best_subject[0], best_subject[2])))

        # Pārbaude vai jauns risinājums ir tabu sarakstā, ja nav pieņemts jauno risinājumu
        # ja ir tad var būt pieņemts ar aspiration kritēju
        if tabu_element not in tabu_list or best_penalty < evaluate_schedule(solution, c)[0]:
            tabu_list.append(tabu_element)
            solution = best_solution
            _, conflicts_list = evaluate_schedule(solution, clases)

            if current_penalty != best_penalty:
                current_penalty = best_penalty
                m_counter = 0
            else:
                m_counter+=1
                
            # Tabu saraksta izmērs
            max_tabu_size = MaxIter * 0.1
            if len(tabu_list) >= max_tabu_size:
                tabu_list.pop(0)
        
    # Ja risinājumam novērtējums ir 0, tad beigt darbību
    if (best_penalty) == 0:
        break
    print(best_penalty)

# Pārbaude: 
#print (check_hard_constraints(solution)) # Jābūt True, atbilst visiem stingriem ierobežojumiem
#print(conflicts_list) # Tukš, kad nav kofliktu - visas prasības ir izpildītas. 
# Ja nav tukš, tad izvada konfliktu pāris, kuri pārkapa prasības 
# Izvada iterāciju skaitu, kas bija nepieciešams, lai atrast galīgo risinājumu
print(itera)

end_time = time.time()
execution_time = end_time - start_time

print(f"Algoritms strādā {execution_time:.4f} sekundes.")

# Excel faila saglābašana
columns = ["Diena", "Laiks"]
for c in clases:
    columns.extend([f"{c} Telpa", f"{c} Skolotājs", f"{c} Priekšmets"])

schedule_df = pd.DataFrame(columns=columns)
dienu_nosaukumi = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturtdiena", "Piektdiena"]

for day_num, diena in enumerate(dienu_nosaukumi, start=1):
    for slot_index, slot in enumerate(solution[day_num], start=1):
        row = {"Diena": diena if slot_index == 1 else "", "Laiks": slot_index}
        for class_index, c in enumerate(clases):
            lesson = slot[class_index]
            if lesson:
                # Ja ir vairākas stundas šajā slotā — kombinētā stunda
                if isinstance(lesson, list):
                    subjects = []
                    teachers = []
                    rooms = []
                    for l in lesson:
                        if len(l) == 3:
                            subj, teach, rm = l
                            subjects.append(str(subj))
                            teachers.append(str(teach))
                            rooms.append(str(rm))
                    row[f"{c} Priekšmets"] = " / ".join(subjects)
                    row[f"{c} Skolotājs"] = " / ".join(teachers)
                    row[f"{c} Telpa"] = " / ".join(rooms)
                elif len(lesson) == 3:
                    subject, teacher, room = lesson
                    row[f"{c} Priekšmets"] = subject
                    row[f"{c} Skolotājs"] = teacher
                    row[f"{c} Telpa"] = room
                else:
                    row[f"{c} Priekšmets"] = ""
                    row[f"{c} Skolotājs"] = ""
                    row[f"{c} Telpa"] = ""
            else:
                row[f"{c} Priekšmets"] = ""
                row[f"{c} Skolotājs"] = ""
                row[f"{c} Telpa"] = ""
        schedule_df = pd.concat([schedule_df, pd.DataFrame([row])], ignore_index=True)
    
    # Tukša rinda starp dienām
    empty_row = {col: "" for col in columns}
    schedule_df = pd.concat([schedule_df, pd.DataFrame([empty_row])], ignore_index=True)

schedule_df.to_excel("Stundu_saraksts_Tabu.xlsx", index=False)
print("Saglabāts fails 'Stundu_saraksts_Tabu.xlsx'")

# Kolonnu auto izlidzināšana pēc teksta
file_path = "Stundu_saraksts_Tabu.xlsx"
schedule_df.to_excel(file_path, index=False)

wb = load_workbook(file_path)
ws = wb.active

for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    
    ws.column_dimensions[col_letter].width = max_length + 2

wb.save(file_path)
