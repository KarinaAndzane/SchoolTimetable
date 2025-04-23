from pulp import LpProblem, LpVariable, lpSum, LpMinimize, LpBinary, GUROBI, value
import math
import pandas as pd
from openpyxl import load_workbook
import re

# Datu lasīšana no Excel faila (klases, skolotāji, telpas, klases skolotāji)
dfKS = pd.read_excel("Skolotaji.xlsx", sheet_name="Klases sk.")
dfS = pd.read_excel("Skolotaji.xlsx", sheet_name="Skolotaji")
dfMP = pd.read_excel("Skolotaji.xlsx", sheet_name="M.priekšmeti")
dfP = pd.read_excel("Skolotaji.xlsx", sheet_name="Programma")

# Programma
dfP.columns = ['Subject', '7', '8', '9']

# Izglītības programmas: priekšmets un stundu skaits nedēļa
programs = {
    "7": dict(zip(dfP['Subject'], dfP['7'])),
    "8": dict(zip(dfP['Subject'], dfP['8'])),
    "9": dict(zip(dfP['Subject'], dfP['9']))
}

# Klase un klases skolotājs
class_teachers = dict(zip(dfKS["Klase"], dfKS["Skolotājs"]))

# Klašu saraksts
classes = list(class_teachers.keys())

# Skolotājs un telpa
rooms = dict(zip(dfS["Skolotājs"], dfS["Telpa"]))

# Mācību priekšmets un skolotāji
teachers = {}

for index, row in dfMP.iterrows():
    subject = row["Mācību priekšmeti"]  
    t_list = [teacher for teacher in row[1:].dropna().tolist() if teacher != subject]  
    teachers[subject] = t_list

teachers["Klases stunda"] = list(class_teachers.values())

# 5 mācību dienas nedēļā
days = range(1, 6) 
day_name = {
    1: "Pirmdiena",
    2: "Otrdiena",
    3: "Trešdiena",
    4: "Ceturtdiena",
    5: "Piektdiena"
} 

# Maksimālais stundu skaits dienā
slots = {"7": 7, "8": 8, "9": 8}  

# Skolotāju noslodze, maksimālais stundu skaits nedēļā
teacher_load = 23


#  SKOLOTĀJU PRASĪBAS

excel_path = "Skolotaji.xlsx"
xlsx = pd.ExcelFile(excel_path)

teachers_df = pd.read_excel(excel_path, sheet_name="Skolotaji")
# Vārdnīca, kur atslēga ir skolotāja numurs, bet vērtība ir V. Uzvārds
teacher_names = dict(zip(teachers_df['Nr.'], teachers_df['Skolotājs']))

# Nepieciešamas Excel lapas 
sheets = [s for s in xlsx.sheet_names if s != 'Klases sk.']

# Prioritātes grupas;
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
    print(teacher_number)
    priority = match.group(2)

    # Meklējam skolotāja vārdu
    if teacher_number in teacher_names:
        full_name = teacher_names[teacher_number]
    else:
        print(f"Skolotājs ar numuru {teacher_number} nav atrast")
        continue

    valid_days = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturtdiena", "Piektdiena"]

    df = pd.read_excel(excel_path, sheet_name=sheet)
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

print(dict1)
print(dict2)


# Visu iespējamo mācību priekšmetu saraksts
subjects = {grade: list(programs[grade[0]].keys()) for grade in classes}

# Pārbuade: Vai pietiks skolotāju, lai novadīt visas stundas, īstenot mācību programmu
for subject, teacher_list in teachers.items():
    total_needed = sum(programs[grade][subject] * len([c for c in classes if c.startswith(grade)]) for grade in programs if subject in programs[grade])
    max_possible = len(teacher_list) * teacher_load
    
    if total_needed > max_possible:
        raise ValueError(f"Nepietiek skolotāju priekšmetam {subject}.Ir nepieciešami {math.ceil(total_needed/teacher_load)} skolotāji, bet pieejami ir {max_possible//teacher_load} skolotāji.")


# MODELIS 
prob = LpProblem("School_Schedule", LpMinimize)

# Mainīgie x[c, s, d, t, tc] = 1, ja priekšmets s klasei c dienā d laika slotā t pasniedz skolotājs tc
x = LpVariable.dicts("x", [(c, s, d, t, tc, rooms[tc]) for c in classes for s in subjects[c] for d in days
                            for t in range(1, slots[c[0]] + 1) for tc in teachers[s]], cat=LpBinary)

# Mainīgie skolotāja piešķiršanai konkrētam priekšmetam klasē                       
y = LpVariable.dicts("y", [(c, s, tc) for c in classes for s in subjects[c] for tc in teachers[s]], cat=LpBinary)

#  Mainīgais, kas norāda, vai laika slots ir aizņemts
z = LpVariable.dicts("z", [(c, d, t) for c in classes for d in days for t in range(1, slots[c[0]] + 1)], cat=LpBinary)

# Mīksto ierobežojumu pievienošana
# Skolotāju vēlami laiki ar prioritāti 2
penalty = lpSum([
    x[(c, s, d, t, tc, rooms[tc])]
    for c in classes
    for s in subjects[c]
    for d in days
    for t in range(1, slots[c[0]] + 1)
    for tc in teachers[s]
    if tc in dict2 and t in dict2[tc].get(day_name[d], [])
])
prob += penalty


# Ierobežojums: skolotāju prasības (1.prioritāte(obligāta)) Vēlami un nevēlami darba laiki
for c in classes:
    for s in subjects[c]:
        for d in days:
            for t in range(1, slots[c[0]] + 1):
                for tc in teachers[s]:
                    if tc in dict1 and t in dict1[tc].get(day_name[d], []):
                        prob += x[(c, s, d, t, tc, rooms[tc])] == 0

# Ierobežojums: klases skolotājam jāvada klases stunda savai klasei
for c, tc in class_teachers.items():
    room =  rooms[tc]
    prob += lpSum(x[c, "Klases stunda", d, t, tc, room]
    for d in days
    for t in range(1, slots[c[0]] + 1)) == programs[c[0]]["Klases stunda"]

# Ierobežojums: katrs priekšmets jāpasniedz noteiktu reižu skaitu nedēļā
for c in classes:
    grade = c[0]
    for s, num_lessons in programs[grade].items():
        prob += lpSum(x[c, s, d, t, tc, rooms[tc]]
        for d in days
        for t in range(1, slots[grade] + 1)
        for tc in teachers[s]) == num_lessons

svesvaloda = ["Vācu valoda", "Franču valoda"]

# Ierobežojums: katrā laika slotā klasei var būt tikai viens priekšmets, izņemot svešvalodas
subjects_without_svesvaloda = {
    grade: [subject for subject in subjects[grade] if subject not in svesvaloda]
    for grade in subjects
}

for c in classes:
    grade = c[0]
    for d in days:
        for t in range(1, slots[grade] + 1):
            prob += lpSum(x[c, s, d, t, tc, rooms[tc]] 
            for s in subjects_without_svesvaloda[c]
            for tc in teachers[s]) <= 1

            prob += lpSum(x[(c, s, d, t, tc, rooms[tc])] 
            for s in subjects[c]
            for tc in teachers[s]) <= 2

# Ierobežojums: ja vienā laika slotā ir vācu valoda, tad tajā pašā slotā jābūt franču valodai
for c in classes:
    for d in days:
        for t in range(1, slots[c[0]] + 1):
            prob += lpSum(x[(c, "Vācu valoda", d, t, tc, rooms[tc])]for tc in teachers["Vācu valoda"]) == \
            lpSum(x[(c, "Franču valoda", d, t, tc, rooms[tc])] for tc in teachers["Franču valoda"])

# Ierobežojums: stundas jāsākas no pirmās stundas
for c in classes:
    for d in days:
        prob += lpSum(x[(c, s, d, 1, tc, rooms[tc])] 
        for s in subjects[c] for tc in teachers[s]) >= 1 # >=1, ja ir svešvaloda, tad x vērtība  laikas slotā t būs vienāds ar 2
        
# Ierobežojums: stundas jābūt bez "logiem" brīvām stundām
for c in classes:
    for d in days:
        for t in range(1, slots[c[0]] + 1):
            sum_lessons = lpSum(
                x[(c, s, d, t, tc, rooms[tc])]
                for s in subjects[c]
                for tc in teachers[s]
            )
            prob += sum_lessons >= z[c, d, t]
            prob += sum_lessons <= z[c, d, t] * 2

# Ierobežojums: ja stundas ir slotā t, tad tās jābūt arī slotā t-1 (ja t >= 2)
for c in classes:
    for d in days:
        for t in range(2, slots[c[0]] + 1):
            prob += z[c, d, t] <= z[c, d, t-1]

# Ierobežojums: viens un tas pats priekšmets nevar atkārtoties vienā dienā (ja būs mīkstie ier. tad būs jāiekļauj izņemumus)
for c in classes:
    grade = c[0]
    for d in days:
        for s in subjects[c]:
            prob += lpSum(x[c, s, d, t, tc, rooms[tc]] 
            for t in range(1, slots[grade] + 1) 
            for tc in teachers[s]) <= 1

# Ierobežojums: stundas jāsadala vienmērīgi pa dienām, uzstadot minimālu stundu skaitu dienā
for c in classes:
    grade = c[0]
    total_lessons = sum(programs[grade].values())  # Kopīgs stundu skaits nedēļā
    min_lessons_per_day = total_lessons // len(days)  # Minimālais stundu skaits dienā
    
    for d in days:
        prob += lpSum(x[c, s, d, t, tc, rooms[tc]]
        for s in subjects[c]
        for t in range(1, slots[grade] + 1)
        for tc in teachers[s]) >= min_lessons_per_day
        
# Ierobežojums: skolotājs vienlaikus var pasniedz tikai vienu stundu
for tc in set([t for sublist in teachers.values() for t in sublist]):
    room = rooms[tc]
    for d in days:
        for t in range(1, max(slots.values()) + 1):
            prob += lpSum(x[c, s, d, t, tc, room] for c in classes for s in subjects[c] 
                if tc in teachers[s] and t <= slots[c[0]]) <= 1

# Ierobežojums: skolotājs nedrīkst parsniegt vairāk nekā 23 stundas nedēļā
for tc in set([t for sublist in teachers.values() for t in sublist]):
    room = rooms[tc]
    prob += lpSum(x[c, s, d, t, tc, room] for c in classes for s in subjects[c] 
                  for d in days for t in range(1, slots[c[0]] + 1)
                  if tc in teachers[s]) <= teacher_load

# Ierobežojums: ja skolotājs pasniedz kaut vienu priekšmeta s stundu klasei c, tad viņam jāpasniedz visas šī priekšmeta stundas šajā klasē
for c in classes:
    for s in subjects[c]:
        for tc in teachers[s]:
            room = rooms[tc]
            # Ja skolotājs ir izvēlēts priekšmeta pasniegšanai, viņam jāvada visas šī priekšmeta stundas šajā klasē
            prob += lpSum(x[c, s, d, t, tc, room] for d in days for t in range(1, slots[c[0]] + 1)) <= y[c, s, tc] * 1000

        # Nodrošinām, ka tikai viens skolotājs pasniedz šo priekšmetu konkrētajai klasei
        prob += lpSum(y[c, s, tc] for tc in teachers[s]) == 1

# Ierobežojums: vienā laika slotā vienu telpu var izmantot tikai viena klase
for room in set(rooms.values()):
    for d in days:
        for t in range(1, max(slots.values()) + 1):
            prob += lpSum(x[c, s, d, t, tc, room] 
                      for c in classes 
                      for s in subjects[c] 
                      for tc in teachers[s] 
                      if rooms[tc] == room and t <= slots[c[0]]) <= 1

# Risinājums
prob.solve(GUROBI())
# Sodu skaits 
print(f"Kopējais sods: {value(penalty)}")

# SKOLOTĀJU DARBA GRAFIKI
all_teachers = set(tc for t_list in teachers.values() for tc in t_list)

teacher_schedule = {tc: {d: {t: [] for t in range(1, max(slots.values()) + 1)} for d in days} for tc in all_teachers}

for c in classes:
    grade = c[0]
    for s in subjects[c]:
        for d in days:
            for t in range(1, slots[grade] + 1):
                for tc in teachers[s]:
                    room = rooms[tc]
                    var_key = (c, s, d, t, tc, room)
                    if var_key in x and x[var_key].value() == 1:
                        teacher_schedule[tc][d][t].append((c, s, room))

# Excel
teacher_columns = ["Diena", "Slots", "Klase", "Priekšmets"]
n = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturdiena", "Piektdiena"]
teacher_schedule_df = pd.DataFrame(columns=["Skolotājs"] + teacher_columns)

for tc in sorted(all_teachers):
    for d in days:
        for t in range(1, max(slots.values()) + 1):
            lessons = teacher_schedule[tc][d][t]
            if lessons:
                for c, s, room in lessons:
                    row = {
                        "Skolotājs": tc,
                        "Diena": n[d - 1],
                        "Slots": t,
                        "Klase": c,
                        "Priekšmets": s,
                    }
                    teacher_schedule_df = pd.concat([teacher_schedule_df, pd.DataFrame([row])], ignore_index=True)

teacher_schedule_df.to_excel("Teacher_Schedule.xlsx", index=False)
print("Saglābāts fails: 'Teacher_Schedule.xlsx'")

# STUNDU SARAKSTA SAGLABĀŠANA EXCEL FAILĀ 
schedule_data = []
columns = ["Dienas", "Slots"]
empty_col_index = 1 

for c in classes:
    columns.extend([f"{c} Telpa", f"{c} Skolotājs", f"{c} Priekšmets"])
    empty_col_index += 1  

schedule_df = pd.DataFrame(columns=columns)

for d in days:
    for t in range(1, max(slots.values()) + 1):
        row_data = {"Dienas": n[d - 1] if t == 1 else "", "Slots": t}

        for i, c in enumerate(classes, start=1):
            if t <= slots[c[0]]:
                subjects_list = []
                teachers_list = []
                rooms_list = []

                for s in subjects[c]:
                    for tc in teachers[s]:
                        room = rooms.get(tc, "")
                        if x[c, s, d, t, tc, room].value() == 1:
                            subjects_list.append(s)
                            teachers_list.append(tc)
                            rooms_list.append(str(room))

                if subjects_list:
                    row_data.update({
                        f"{c} Priekšmets": " / ".join(subjects_list),
                        f"{c} Skolotājs": " / ".join(teachers_list),
                        f"{c} Telpa": " / ".join(rooms_list),
                    })
                else:  # Ja nav stundas
                    row_data.update({
                        f"{c} Priekšmets": "",
                        f"{c} Skolotājs": "",
                        f"{c} Telpa": "",
                    })
            else:  # Maksimāls stundu skaits, dienas slots
                row_data.update({
                    f"{c} Priekšmets": "",
                    f"{c} Skolotājs": "",
                    f"{c} Telpa": "",
                })

        schedule_df = pd.concat([schedule_df, pd.DataFrame([row_data])], ignore_index=True)

    empty_row = {col: "" for col in columns}
    schedule_df = pd.concat([schedule_df, pd.DataFrame([empty_row])], ignore_index=True)

schedule_df.to_excel("Stundu_saraksts.xlsx", index=False)

print("Saglabāts fails 'Stundu_saraksts.xlsx'")


# Kolonnu auto izlidzināšana pēc teksta
file_path = "Stundu_saraksts.xlsx"
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



