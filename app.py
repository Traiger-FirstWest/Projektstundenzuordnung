import streamlit as st
from datetime import date
import pandas as pd
import datetime
import random
import copy
from io import BytesIO
import holidays

def assign_hours(rows, active_tasks, mode: str, min_hours: float, max_hours: float,
                 earliest_hour: int = 9, latest_hour: int = 13):
    """
    rows: Liste wie bisher (header + Datenzeilen)
    active_tasks: {emp: [ {project_id,start,end,remaining_hours, ...}, ... ]}
    mode:
      - "forward"  (Anfang -> Ende)
      - "backward" (Ende -> Anfang)
      - "flat"     (flach über Zeitraum)
    """

    header = rows[0]
    # Spaltenindizes (wie bei dir)
    idx_name = 0
    idx_date = 1
    idx_time = 3
    idx_info = 4
    idx_hours = 5

    # Helfer: parse datum
    def parse_date(s: str):
        return datetime.datetime.strptime(s, "%d.%m.%Y").date()

    # Helfer: zufällige Zeit
    def rand_time(d: datetime.date):
        hour = random.randint(earliest_hour, max(earliest_hour, latest_hour - 1))
        minute = random.randint(0, 59)
        start_dt = datetime.datetime.combine(d, datetime.time(hour=hour, minute=minute))
        return start_dt

    # Precompute: rows je Mitarbeiter (nur "anwesend")
    work_rows_by_emp = {emp: [] for emp in active_tasks.keys()}
    for r_i in range(1, len(rows)):
        row = rows[r_i]
        emp = row[idx_name]
        if emp not in active_tasks:
            continue
        if row[idx_info] != "anwesend":
            continue
        try:
            d = parse_date(row[idx_date])
        except Exception:
            continue
        work_rows_by_emp[emp].append((d, r_i))

    # Reihenfolge der Tage pro Mitarbeiter
    for emp, items in work_rows_by_emp.items():
        items.sort(key=lambda x: x[0])  # chronologisch
        if mode == "backward":
            items.reverse()
        work_rows_by_emp[emp] = items

    # Für "flat": Track pro Task, wie viel schon gebucht wurde
    # wir hängen Felder an active_tasks dran (assigned_so_far, total_hours)
    if mode == "flat":
        for emp, tasks in active_tasks.items():
            for t in tasks:
                total = float(t.get("remaining_hours", 0.0))
                t["total_hours"] = total
                t["assigned_so_far"] = 0.0

    # Buchung pro Arbeitstag: max 1
    for emp, day_rows in work_rows_by_emp.items():
        tasks = active_tasks.get(emp, [])
        if not tasks:
            continue

        for d, r_i in day_rows:
            # wenn emp nichts mehr offen hat, abbrechen
            if sum(t.get("remaining_hours", 0.0) for t in tasks) <= 0:
                break

            # passende Tasks (im Datumsfenster + remaining > 0)
            eligible = [t for t in tasks if t["start"] <= d <= t["end"] and t["remaining_hours"] > 0]
            if not eligible:
                continue

            # Task auswählen je nach Modus
            if mode in ("forward", "backward"):
                # simpel: nimm den Task mit den meisten Reststunden (oder random)
                eligible.sort(key=lambda t: t["remaining_hours"], reverse=True)
                chosen = eligible[0]

            elif mode == "flat":
                # wähle Task, der "hinter" seinem Soll liegt (Backlog am größten)
                best = None
                best_score = None

                for t in eligible:
                    total = float(t.get("total_hours", 0.0)) or (float(t["remaining_hours"]) + float(t.get("assigned_so_far", 0.0)))
                    assigned = float(t.get("assigned_so_far", 0.0))
                    # Fortschritt im Task-Zeitraum
                    span = (t["end"] - t["start"]).days
                    if span <= 0:
                        frac = 1.0
                    else:
                        frac = (d - t["start"]).days / span
                        if frac < 0:
                            frac = 0.0
                        if frac > 1:
                            frac = 1.0

                    expected = total * frac
                    backlog = expected - assigned  # >0 => hinter Plan

                    # Wir wählen größten backlog; wenn alle negativ, nimmt der "am meisten remaining"
                    score = backlog
                    if best is None or score > best_score:
                        best = t
                        best_score = score

                chosen = best if best is not None else eligible[0]

            else:
                chosen = eligible[0]

            # Stunden-Dauer bestimmen
            duration = round(random.uniform(min_hours, max_hours), 1)
            duration = min(duration, float(chosen["remaining_hours"]))
            if duration <= 0:
                continue

            # Buchung schreiben
            start_dt = rand_time(d)
            end_dt = start_dt + datetime.timedelta(hours=float(duration))
            rows[r_i][idx_time] = f"{start_dt.strftime('%H:%M')}-{end_dt.strftime('%H:%M')}"
            rows[r_i][idx_info] = chosen["project_id"]     # wie in deinem Code: Info wird Projektcode
            rows[r_i][idx_hours] = float(round(duration, 2))

            chosen["remaining_hours"] -= float(duration)
            if mode == "flat":
                chosen["assigned_so_far"] = float(chosen.get("assigned_so_far", 0.0)) + float(duration)


### functions
def generate_random_start_time(base_date, earliest_hour=9, latest_hour=13):
    hour = random.randint(earliest_hour, max(earliest_hour, latest_hour - 1))
    minute = random.randint(0, 59)
    return datetime.datetime.combine(base_date, datetime.time(hour=hour, minute=minute))


def prepare_active_tasks(tasks_per_employee):
    active_tasks = {}
    for emp, task_list in tasks_per_employee.items():
        active_tasks[emp] = []
        for task in task_list:
            t = copy.deepcopy(task)
            t["remaining_hours"] = float(t.pop("hours", 0.0))
            active_tasks[emp].append(t)
    return active_tasks


def assign_one_task_per_day(rows, active_tasks, project_id_col=4, time_col=3, min_hours=2.0, max_hours=6.0):
    hours_col = 5  # "Stunden"
    for row in rows[1:]:  # header überspringen
        name = row[0]
        datum_str = row[1]
        info = row[project_id_col]

        if info != "anwesend" or name not in active_tasks:
            continue

        try:
            datum = datetime.datetime.strptime(datum_str, "%d.%m.%Y").date()
        except ValueError:
            continue

        tasks = active_tasks[name]
        random.shuffle(tasks)

        for task in tasks:
            if task["start"] <= datum <= task["end"] and task["remaining_hours"] > 0:
                duration = round(random.uniform(min_hours, max_hours), 1)
                duration = min(duration, task["remaining_hours"])

                start_time = generate_random_start_time(datum)
                end_time = start_time + datetime.timedelta(hours=float(duration))
                time_str = f"{start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')}"

                row[time_col] = time_str
                row[project_id_col] = task["project_id"]
                row[hours_col] = float(round(duration, 2))

                task["remaining_hours"] -= float(duration)
                break  # max 1 task pro Tag


def build_project_mappings(project_tree):
    """
    Baut:
    - key_to_code: map von internem Task-Key (FSid_APid_Tid) -> Code FSx_APy_Tz
    - legend_rows: Liste für Legenden-Sheet
    - tasks_meta: flache Liste aller Tätigkeiten mit FS/AP/T + Code + internem Key
    """
    key_to_code = {}
    legend_rows = []
    tasks_meta = []

    for fs_i, fs in enumerate(project_tree, start=1):
        fs_code = f"FS{fs_i}"
        for ap_i, ap in enumerate(fs.get("aps", []), start=1):
            ap_code = f"{fs_code}_AP{ap_i}"
            for t_i, t in enumerate(ap.get("tasks", []), start=1):
                code = f"{ap_code}_T{t_i}"
                internal_key = f"{fs['id']}_{ap['id']}_{t['id']}"
                key_to_code[internal_key] = code

                legend_rows.append({
                    "Code": code,
                    "Forschungsschwerpunkt": fs["name"],
                    "Arbeitspaket": ap["name"],
                    "Tätigkeit": t["name"],
                })

                tasks_meta.append({
                    "fs": fs, "ap": ap, "t": t,
                    "code": code,
                    "internal_key": internal_key
                })

    return key_to_code, legend_rows, tasks_meta


def build_tasks_per_employee(project_tree, hours_by_task):
    """
    Aus deinem Streamlit-Baum + Stundenmapping wird die tasks_per_employee Struktur:
    { employee: [ {project_id,start,end,hours,fs,ap,t}, ... ] }
    """
    key_to_code, legend_rows, tasks_meta = build_project_mappings(project_tree)
    tasks_per_employee = {}

    for item in tasks_meta:
        fs = item["fs"]
        ap = item["ap"]
        t = item["t"]
        code = item["code"]
        k = item["internal_key"]

        fs_start = fs.get("start")
        fs_end = fs.get("end")
        if fs_start is None or fs_end is None:
            continue

        emp_map = hours_by_task.get(k, {})
        for emp, hrs in emp_map.items():
            try:
                h = float(hrs)
            except Exception:
                h = 0.0
            if h <= 0:
                continue

            entry = {
                "project_id": code,
                "fs": fs["name"],
                "ap": ap["name"],
                "t": t["name"],
                "start": fs_start,
                "end": fs_end,
                "hours": h
            }
            tasks_per_employee.setdefault(emp, []).append(entry)

    return tasks_per_employee, legend_rows


def build_base_rows(employees, wj_start, wj_end, absences_by_employee):
    header = ["Name", "Datum", "Wochentag", "Zeitbuchung", "Info", "Stunden"]
    rows = [header]

    # Feiertage (AT + Wien default)
    years = list(range(wj_start.year, wj_end.year + 1))
    at_holidays = holidays.country_holidays("AT", years=years)

    date_list = []
    d = wj_start
    while d <= wj_end:
        date_list.append(d)
        d += datetime.timedelta(days=1)

    for name in employees:
        abs_list = absences_by_employee.get(name, [])
        for datum in date_list:
            datum_str = datum.strftime("%d.%m.%Y")
            wochentag = datum.strftime("%A")

            if datum.weekday() in (5, 6):
                info = "Wochenende"
            elif datum in at_holidays:
                info = at_holidays.get(datum)
            else:
                info = "anwesend"
                for e in abs_list:
                    if e["start"] <= datum <= e["end"]:
                        info = e["reason"]
                        break

            rows.append([name, datum_str, wochentag, "", info, ""])

    return rows

STEPS = [
    "Start",
    "Wirtschaftsjahr",
    "Mitarbeiter",
    "Abwesenheiten",
    "Projekte",
    "Stundenzuordnung",
    "Export",
]

def init_state():
    if "step" not in st.session_state:
        st.session_state.step = 0

    # Datencontainer für spätere Schritte - Teil 1
    st.session_state.setdefault("wirtschaft_start", None)
    st.session_state.setdefault("wirtschaft_end", None)

    # Für die Mitarbeiterlisten - Teil 2
    st.session_state.setdefault("employees_text", "")
    st.session_state.setdefault("employees_df", pd.DataFrame({"Name": []}))
    st.session_state.setdefault("employees", [])

    #Abwesenhetien - Teil 3
    st.session_state.setdefault("absence_employee_idx", 0)          # welcher Mitarbeiter gerade dran ist
    st.session_state.setdefault("absences_by_employee", {})         # dict: { "Name": [ {start,end,reason}, ... ] }
    st.session_state.setdefault("absences_done", set())             # set der Mitarbeiter, die "fertig" bestätigt wurden

    ### Teil 4 (?)
    # Projektbaum
    st.session_state.setdefault("project_tree", [])  # Liste von FS
    st.session_state.setdefault("proj_sel", {"kind": None, "fs": None, "ap": None, "t": None})

    # einfache Zähler für stabile IDs
    st.session_state.setdefault("fs_counter", 0)
    st.session_state.setdefault("ap_counter", 0)
    st.session_state.setdefault("t_counter", 0)

    st.session_state.setdefault("hours_sel", {"fs": None, "ap": None, "t": None})
    st.session_state.setdefault("hours_by_task", {})  # key -> {employee: hours}               
    ###                  

def go_next():
    st.session_state.step = min(st.session_state.step + 1, len(STEPS) - 1)

def go_prev():
    st.session_state.step = max(st.session_state.step - 1, 0)

def sidebar():
    st.sidebar.title("Übersicht")
    for i, name in enumerate(STEPS):
        if i == st.session_state.step:
            st.sidebar.markdown(f"➡️ **{i+1}. {name}**")
        elif i < st.session_state.step:
            st.sidebar.markdown(f"✅ {i+1}. {name}")
        else:
            st.sidebar.markdown(f"• {i+1}. {name}")

    st.sidebar.divider()
    st.sidebar.progress(st.session_state.step / (len(STEPS) - 1))

def page(title: str, show_prev=True, show_next=True, next_disabled=False):
    st.header(title)

    col1, col2 = st.columns(2)
    with col1:
        st.button("← Zurück", on_click=go_prev, disabled=not show_prev)
    with col2:
        st.button("Weiter →", on_click=go_next, disabled=(not show_next) or next_disabled)

def main():
    st.set_page_config(page_title="Streamlit Wizard", layout="wide")
    init_state()
    sidebar()

    step = st.session_state.step
#####################################################################################
    if step == 0:
        # Logo-Zeile (links), rechts bewusst leer = "in der Ecke"
        col_logo, col_spacer = st.columns([2, 8], vertical_alignment="top")
        with col_logo:
            try:
                st.image("assets/logo.png", width=320)  # hier größer/kleiner machen
            except Exception:
                pass

        st.title("Projektstundenzuordnung – Excel Generator")

        st.markdown(
            """
            Dieses Tool unterstützt dich dabei, **Projektstunden je Mitarbeiter** zu erfassen und daraus eine **Excel-Datei** zu erzeugen.

            **Ablauf:**
            1. Wirtschaftsjahr festlegen  
            2. Mitarbeiter erfassen  
            3. Abwesenheiten eintragen  
            4. Projektstruktur (FS/AP/Tätigkeiten) definieren  
            5. Stunden je Tätigkeit & Mitarbeiter zuordnen  
            6. Excel exportieren  

            **Hinweis:** Bitte prüfe das Ergebnis (Plausibilität), bevor du es weiterverwendest.
            """
        )

        st.divider()
        st.caption("Autor: Emanuel Traiger")

        ack = st.checkbox("Verstanden – ich möchte starten", value=False)

        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev, disabled=True)
        with col2:
            st.button("Weiter →", on_click=go_next, disabled=not ack)
#########################################################################################
    elif step == 1:
        st.header("Wirtschaftsjahr")
        st.write("Bitte wähle den Zeitraum, für den die Zeiterfassung generiert werden soll.")

        colA, colB = st.columns(2)

        with colA:
            start = st.date_input(
                "Startdatum",
                value=st.session_state.wirtschaft_start or date(date.today().year, 1, 1),
                key="wirtschaft_start_input",
            )

        with colB:
            end = st.date_input(
                "Enddatum",
                value=st.session_state.wirtschaft_end or date(date.today().year, 12, 31),
                key="wirtschaft_end_input",
            )

        # Speichern in session_state (damit es später verfügbar ist)
        st.session_state.wirtschaft_start = start
        st.session_state.wirtschaft_end = end

        # Validierung
        valid = start < end

        if not valid:
            st.error("Das Startdatum muss vor dem Enddatum liegen.")
 
        st.divider()
        st.caption(f"Aktuell gewählt: {st.session_state.wirtschaft_start} bis {st.session_state.wirtschaft_end}")

        # Navigation: Weiter nur wenn valid
        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev)
        with col2:
            st.button("Weiter →", on_click=go_next, disabled=not valid)
#########################################################################################
    elif step == 2:
        st.header("Mitarbeiter")
        st.write("Füge hier deine Mitarbeiter ein. **Eine Zeile = ein Name**. Du kannst auch eine Spalte aus Excel kopieren und einfügen.")

        # --- 1) Reinkopieren / Eingeben
        st.text_area(
            "Mitarbeiterliste",
            key="employees_text",
            height=180,
            placeholder="Max Mustermann\nErika Musterfrau\n..."
        )

        def parse_names(text: str) -> list[str]:
            names = []
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                # Falls aus Excel kopiert: manchmal sind Tabs drin (mehrere Spalten)
                line = line.split("\t")[0].strip()
                if line:
                    names.append(line)
            return names

        colA, colB = st.columns([1, 2])
        with colA:
            if st.button("Übernehmen"):
                names = parse_names(st.session_state.employees_text)
                # In Tabelle speichern (für Editieren)
                st.session_state.employees_df = pd.DataFrame({"Name": names})

        with colB:
            st.caption("Tipp: Nach „Übernehmen“ kannst du unten in der Tabelle Namen korrigieren oder neue Zeilen hinzufügen.")

        st.divider()

        # --- 2) Tabelle zum Nachbearbeiten
        st.subheader("Liste (bearbeitbar)")
        edited_df = st.data_editor(
            st.session_state.employees_df,
            num_rows="dynamic",
            use_container_width=True,
            key="employees_editor",
        )

        # Clean & Validierung
        cleaned = []
        for x in edited_df.get("Name", []):
            if x is None:
                continue
            name = str(x).strip()
            if name:
                cleaned.append(name)

        # Duplikate checken (case-insensitive)
        lower = [n.lower() for n in cleaned]
        duplicates = sorted({n for n in cleaned if lower.count(n.lower()) > 1})

        if len(cleaned) == 0:
            st.warning("Bitte mindestens einen Mitarbeiter eintragen.")
        if duplicates:
            st.error(f"Duplikate gefunden (bitte bereinigen): {', '.join(duplicates)}")

        # Speichern für nächste Schritte
        st.session_state.employees = cleaned
        st.session_state.employees_df = pd.DataFrame({"Name": cleaned})

        # Navigation: Weiter nur wenn valide & eindeutig
        can_continue = (len(cleaned) > 0) and (len(duplicates) == 0)

        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev)
        with col2:
            st.button("Weiter →", on_click=go_next, disabled=not can_continue)
#########################################################################################
    elif step == 3:
        st.header("Abwesenheiten")

        employees = st.session_state.get("employees", [])
        wj_start = st.session_state.get("wirtschaft_start", None)
        wj_end = st.session_state.get("wirtschaft_end", None)

        if not employees:
            st.warning("Bitte zuerst Mitarbeiter anlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        if wj_start is None or wj_end is None:
            st.warning("Bitte zuerst das Wirtschaftsjahr festlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        # --- State aufräumen / initialisieren ---
        # Stelle sicher, dass es für jeden aktuellen Mitarbeiter einen Eintrag gibt
        absences_by_employee = st.session_state.absences_by_employee
        for emp in employees:
            absences_by_employee.setdefault(emp, [])
        # Entferne evtl. alte Mitarbeiter, die nicht mehr in der Liste sind
        for emp in list(absences_by_employee.keys()):
            if emp not in employees:
                del absences_by_employee[emp]
                if emp in st.session_state.absences_done:
                    st.session_state.absences_done.remove(emp)

        # Mitarbeiterindex absichern
        st.session_state.absence_employee_idx = max(
            0, min(st.session_state.absence_employee_idx, len(employees) - 1)
        )

        # --- Fortschritt ---
        done_count = len(st.session_state.absences_done)
        st.caption(f"Fortschritt: {done_count} von {len(employees)} Mitarbeitern bestätigt")
        st.progress(done_count / len(employees) if employees else 0)

        st.divider()

        # --- Mitarbeiter Navigation ---
        idx = st.session_state.absence_employee_idx
        current_emp = employees[idx]

        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            st.button("← Vorheriger", disabled=(idx == 0),
                    on_click=lambda: st.session_state.update({"absence_employee_idx": idx - 1}))
        with col2:
            chosen = st.selectbox(
                "Mitarbeiter",
                options=employees,
                index=idx,
            )
            # wenn Dropdown geändert, Index anpassen
            st.session_state.absence_employee_idx = employees.index(chosen)
            current_emp = chosen
            idx = st.session_state.absence_employee_idx
        with col3:
            st.button("Nächster →", disabled=(idx == len(employees) - 1),
                    on_click=lambda: st.session_state.update({"absence_employee_idx": idx + 1}))

        is_done = current_emp in st.session_state.absences_done
        if is_done:
            st.success("Dieser Mitarbeiter ist als erledigt markiert ✅")

        st.subheader(f"Abwesenheiten für: {current_emp}")

        # --- Abwesenheit hinzufügen (Form) ---
        with st.form("add_absence", clear_on_submit=False):
            # Default = heutiger Tag im WJ, sonst WJ-Start
            default_day = wj_start

            date_range = st.date_input(
                "Zeitraum (von–bis)",
                value=(default_day, default_day),
                min_value=wj_start,
                max_value=wj_end,
            )

            reason = st.selectbox(
                "Grund",
                ["Urlaub", "Krankenstand", "Arztbesuch", "Sonstige"],
            )

            add = st.form_submit_button("Hinzufügen")

        def normalize_range(r):
            # st.date_input kann date oder tuple liefern
            if isinstance(r, tuple) or isinstance(r, list):
                if len(r) == 2:
                    return r[0], r[1]
            return r, r

        if add:
            start, end = normalize_range(date_range)

            if start is None or end is None:
                st.error("Bitte einen gültigen Zeitraum wählen.")
            elif start > end:
                st.error("Startdatum muss vor Enddatum liegen.")
            elif start < wj_start or end > wj_end:
                st.error("Der Zeitraum muss innerhalb des Wirtschaftsjahres liegen.")
            else:
                absences_by_employee[current_emp].append(
                    {"start": start, "end": end, "reason": reason}
                )
                st.session_state.absences_by_employee = absences_by_employee
                # optional: sofort als done markieren, wenn man etwas eingetragen hat
                st.session_state.absences_done.add(current_emp)
                st.rerun()

        st.divider()

        # --- Liste anzeigen + Löschen ---
        entries = absences_by_employee.get(current_emp, [])

        if not entries:
            st.info("Noch keine Abwesenheiten eingetragen.")
        else:
            st.write("Eingetragene Abwesenheiten:")
            for i, e in enumerate(entries):
                cA, cB, cC, cD = st.columns([2, 2, 2, 1])
                with cA:
                    st.write(f"**Von:** {e['start']}")
                with cB:
                    st.write(f"**Bis:** {e['end']}")
                with cC:
                    st.write(f"**Grund:** {e['reason']}")
                with cD:
                    if st.button("Löschen", key=f"del_{current_emp}_{i}"):
                        absences_by_employee[current_emp].pop(i)
                        st.session_state.absences_by_employee = absences_by_employee
                        st.rerun()

        st.divider()

        # --- "Keine Abwesenheiten" bestätigen ---
        cL, cR = st.columns([2, 1])
        with cL:
            st.button(
                "Keine Abwesenheiten (für diesen Mitarbeiter) bestätigen",
                on_click=lambda: st.session_state.absences_done.add(current_emp),
            )
            if is_done:
                st.button(
                    "Markierung zurücksetzen",
                    on_click=lambda: st.session_state.absences_done.discard(current_emp),
                )

        # --- Navigation unten: Weiter nur wenn alle bestätigt ---
        all_done = len(st.session_state.absences_done) == len(employees)

        with cR:
            st.button("← Zurück", on_click=go_prev)

            st.button("Weiter →", on_click=go_next, disabled=not all_done)

        if not all_done:
            st.caption("Hinweis: Bitte jeden Mitarbeiter einmal bestätigen (auch wenn keine Abwesenheiten vorliegen).")
#########################################################################################
    elif step == 4:
        st.header("Projekte")

        wj_start = st.session_state.get("wirtschaft_start", None)
        wj_end = st.session_state.get("wirtschaft_end", None)

        if wj_start is None or wj_end is None:
            st.warning("Bitte zuerst das Wirtschaftsjahr festlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        tree = st.session_state.project_tree
        sel = st.session_state.proj_sel

        def new_id(prefix: str) -> str:
            if prefix == "FS":
                st.session_state.fs_counter += 1
                return f"FS{st.session_state.fs_counter}"
            if prefix == "AP":
                st.session_state.ap_counter += 1
                return f"AP{st.session_state.ap_counter}"
            if prefix == "T":
                st.session_state.t_counter += 1
                return f"T{st.session_state.t_counter}"
            raise ValueError("unknown prefix")

        def count_tasks() -> int:
            c = 0
            for fs in tree:
                for ap in fs["aps"]:
                    c += len(ap["tasks"])
            return c

        def fs_is_complete(fs) -> bool:
            # Minimal-Definition: Zeitraum gesetzt + mind. 1 Tätigkeit irgendwo
            has_window = fs.get("start") is not None and fs.get("end") is not None
            has_any_task = any(len(ap["tasks"]) > 0 for ap in fs["aps"])
            return has_window and has_any_task

        # -----------------------------
        # Layout: links Tree, rechts Details
        # -----------------------------
        left, right = st.columns([1, 2], vertical_alignment="top")

        with left:
            st.subheader("Übersichtsbaum")

            # global: FS hinzufügen
            if st.button("➕ Forschungsschwerpunkt hinzufügen"):
                tree.append({
                    "id": new_id("FS"),
                    "name": f"Neuer Forschungsschwerpunkt {len(tree)+1}",
                    "start": wj_start,
                    "end": wj_end,
                    "aps": []
                })
                st.session_state.project_tree = tree
                st.session_state.proj_sel = {"kind": "fs", "fs": len(tree)-1, "ap": None, "t": None}
                st.rerun()

            st.caption("🔵 FS  |  ⚪️ AP  |  🔸 Tätigkeit")

            if not tree:
                st.info("Noch keine Forschungsschwerpunkte angelegt.")
            else:
                # Tree klickbar rendern
                for i_fs, fs in enumerate(tree):
                    status = "✅" if fs_is_complete(fs) else "⚠️"
                    if st.button(f"{status} 🔵 {fs['name']}", key=f"fs_btn_{fs['id']}"):
                        st.session_state.proj_sel = {"kind": "fs", "fs": i_fs, "ap": None, "t": None}
                        st.rerun()

                    # APs
                    for i_ap, ap in enumerate(fs["aps"]):
                        ap_status = "✅" if len(ap["tasks"]) > 0 else "⚠️"
                        if st.button(f"└─ {ap_status} ⚪️ {ap['name']}", key=f"ap_btn_{fs['id']}_{ap['id']}"):
                            st.session_state.proj_sel = {"kind": "ap", "fs": i_fs, "ap": i_ap, "t": None}
                            st.rerun()

                        # Tasks
                        for i_t, t in enumerate(ap["tasks"]):
                            if st.button(f"   └─ 🔸 {t['name']}", key=f"t_btn_{fs['id']}_{ap['id']}_{t['id']}"):
                                st.session_state.proj_sel = {"kind": "t", "fs": i_fs, "ap": i_ap, "t": i_t}
                                st.rerun()

            st.divider()
            st.caption(f"Summe Tätigkeiten: {count_tasks()}")

        with right:
            st.subheader("Details")

            if sel["kind"] is None:
                st.info("Wähle links einen Forschungsschwerpunkt, ein Arbeitspaket oder eine Tätigkeit aus.")
            else:
                kind = sel["kind"]

                # -----------------------------
                # Forschungsschwerpunkt Details
                # -----------------------------
                if kind == "fs":
                    fs = tree[sel["fs"]]

                    st.markdown("### 🔵 Forschungsschwerpunkt")

                    with st.form("fs_form"):
                        name = st.text_input("Name", value=fs["name"])
                        c1, c2 = st.columns(2)
                        with c1:
                            start = st.date_input("Start", value=fs.get("start", wj_start), min_value=wj_start, max_value=wj_end)
                        with c2:
                            end = st.date_input("Ende", value=fs.get("end", wj_end), min_value=wj_start, max_value=wj_end)

                        save = st.form_submit_button("Speichern")

                    if save:
                        if start > end:
                            st.error("Start muss vor Ende liegen.")
                        else:
                            fs["name"] = name.strip() if name.strip() else fs["name"]
                            fs["start"] = start
                            fs["end"] = end
                            st.session_state.project_tree = tree
                            st.success("Gespeichert.")
                            st.rerun()

                    colA, colB, colC = st.columns([1, 1, 2])
                    with colA:
                        if st.button("➕ Arbeitspaket hinzufügen"):
                            fs["aps"].append({
                                "id": new_id("AP"),
                                "name": f"Neues AP {len(fs['aps'])+1}",
                                "tasks": []
                            })
                            st.session_state.project_tree = tree
                            st.session_state.proj_sel = {"kind": "ap", "fs": sel["fs"], "ap": len(fs["aps"])-1, "t": None}
                            st.rerun()

                    with colB:
                        if st.button("🗑️ Löschen", type="secondary"):
                            # FS löschen
                            tree.pop(sel["fs"])
                            st.session_state.project_tree = tree
                            st.session_state.proj_sel = {"kind": None, "fs": None, "ap": None, "t": None}
                            st.rerun()

                    st.caption("Tipp: Ein FS gilt als „fertig“, wenn Zeitraum gesetzt ist und mind. eine Tätigkeit existiert.")

                # -----------------------------
                # Arbeitspaket Details
                # -----------------------------
                elif kind == "ap":
                    fs = tree[sel["fs"]]
                    ap = fs["aps"][sel["ap"]]

                    st.markdown("### ⚪️ Arbeitspaket")
                    st.caption(f"Im Forschungsschwerpunkt: {fs['name']}")

                    with st.form("ap_form"):
                        name = st.text_input("Name", value=ap["name"])
                        save = st.form_submit_button("Speichern")

                    if save:
                        ap["name"] = name.strip() if name.strip() else ap["name"]
                        st.session_state.project_tree = tree
                        st.success("Gespeichert.")
                        st.rerun()

                    colA, colB, colC = st.columns([1, 1, 2])
                    with colA:
                        if st.button("➕ Tätigkeit hinzufügen"):
                            ap["tasks"].append({
                                "id": new_id("T"),
                                "name": f"Neue Tätigkeit {len(ap['tasks'])+1}",
                            })
                            st.session_state.project_tree = tree
                            st.session_state.proj_sel = {"kind": "t", "fs": sel["fs"], "ap": sel["ap"], "t": len(ap["tasks"])-1}
                            st.rerun()

                    with colB:
                        if st.button("🗑️ Löschen", type="secondary"):
                            fs["aps"].pop(sel["ap"])
                            st.session_state.project_tree = tree
                            st.session_state.proj_sel = {"kind": "fs", "fs": sel["fs"], "ap": None, "t": None}
                            st.rerun()

                    st.caption("Tipp: Ein AP gilt als „fertig“, sobald mindestens eine Tätigkeit darunter existiert.")

                # -----------------------------
                # Tätigkeit Details
                # -----------------------------
                elif kind == "t":
                    fs = tree[sel["fs"]]
                    ap = fs["aps"][sel["ap"]]
                    t = ap["tasks"][sel["t"]]

                    st.markdown("### 🔸 Tätigkeit")
                    st.caption(f"FS: {fs['name']}  |  AP: {ap['name']}")

                    with st.form("t_form"):
                        name = st.text_input("Name", value=t["name"])
                        save = st.form_submit_button("Speichern")

                    if save:
                        t["name"] = name.strip() if name.strip() else t["name"]
                        st.session_state.project_tree = tree
                        st.success("Gespeichert.")
                        st.rerun()

                    colA, colB, colC = st.columns([1, 1, 2])
                    with colB:
                        if st.button("🗑️ Löschen", type="secondary"):
                            ap["tasks"].pop(sel["t"])
                            st.session_state.project_tree = tree
                            st.session_state.proj_sel = {"kind": "ap", "fs": sel["fs"], "ap": sel["ap"], "t": None}
                            st.rerun()

        st.divider()

        # Navigation unten: Weiter erst wenn mind. 1 Tätigkeit existiert
        can_continue = count_tasks() > 0

        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev)
        with col2:
            st.button("Weiter →", on_click=go_next, disabled=not can_continue)

        if not can_continue:
            st.caption("Bitte mindestens eine Tätigkeit anlegen, bevor du fortfährst.")
#########################################################################################
    elif step == 5:
        st.header("Stundenzuordnung der Mitarbeiter auf die Projekte")
        st.write("Wähle links eine Tätigkeit aus und ordne rechts pro Mitarbeiter Stunden zu. "
                "✅ = Stunden vorhanden, ⚠️ = noch keine Stunden.")

        employees = st.session_state.get("employees", [])
        tree = st.session_state.get("project_tree", [])

        if not employees:
            st.warning("Bitte zuerst Mitarbeiter anlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        # --- Hilfsfunktionen ---
        def task_key(fs, ap, t) -> str:
            return f"{fs['id']}_{ap['id']}_{t['id']}"

        def flatten_tasks():
            items = []
            for i_fs, fs in enumerate(tree):
                for i_ap, ap in enumerate(fs.get("aps", [])):
                    for i_t, t in enumerate(ap.get("tasks", [])):
                        items.append({
                            "i_fs": i_fs, "i_ap": i_ap, "i_t": i_t,
                            "fs": fs, "ap": ap, "t": t,
                            "key": task_key(fs, ap, t)
                        })
            return items

        def sum_hours_for_key(k: str) -> float:
            m = st.session_state.hours_by_task.get(k, {})
            total = 0.0
            for v in m.values():
                try:
                    total += float(v)
                except Exception:
                    pass
            return total

        tasks = flatten_tasks()
        if not tasks:
            st.warning("Bitte zuerst im Schritt „Projekte“ mindestens eine Tätigkeit anlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        # --- Aufräumen: entferne alte Task-Keys / alte Mitarbeiter ---
        valid_keys = {it["key"] for it in tasks}
        for k in list(st.session_state.hours_by_task.keys()):
            if k not in valid_keys:
                del st.session_state.hours_by_task[k]
        for k, m in st.session_state.hours_by_task.items():
            for emp in list(m.keys()):
                if emp not in employees:
                    del m[emp]

        # --- Auswahl initialisieren / absichern ---
        sel = st.session_state.hours_sel
        if sel["fs"] is None or sel["ap"] is None or sel["t"] is None:
            first = tasks[0]
            st.session_state.hours_sel = {"fs": first["i_fs"], "ap": first["i_ap"], "t": first["i_t"]}
            sel = st.session_state.hours_sel

        # Falls Auswahl ungültig geworden ist (z.B. nach Löschen)
        def get_selected_task():
            for it in tasks:
                if it["i_fs"] == sel["fs"] and it["i_ap"] == sel["ap"] and it["i_t"] == sel["t"]:
                    return it
            return tasks[0]

        selected = get_selected_task()
        st.session_state.hours_sel = {"fs": selected["i_fs"], "ap": selected["i_ap"], "t": selected["i_t"]}

        # --- Fortschritt ---
        done = sum(1 for it in tasks if sum_hours_for_key(it["key"]) > 0)
        total = len(tasks)

        st.caption(f"Fortschritt: {done} / {total} Tätigkeiten haben Stunden")
        st.progress(done / total if total else 0)

        st.divider()

        # -----------------------------
        # Layout: links Baum, rechts Tabelle
        # -----------------------------
        left, right = st.columns([1, 2], vertical_alignment="top")

        with left:
            st.subheader("Übersichtsbaum")

            search = st.text_input("Suche", placeholder="FS / AP / Tätigkeit…")

            # Nächste offene Tätigkeit (⚠️)
            def jump_next_open():
                # lineare Reihenfolge, ab aktueller Position
                cur_idx = 0
                for idx, it in enumerate(tasks):
                    if it["key"] == selected["key"]:
                        cur_idx = idx
                        break
                order = list(range(cur_idx + 1, len(tasks))) + list(range(0, cur_idx + 1))
                for idx in order:
                    it = tasks[idx]
                    if sum_hours_for_key(it["key"]) <= 0:
                        st.session_state.hours_sel = {"fs": it["i_fs"], "ap": it["i_ap"], "t": it["i_t"]}
                        return True
                return False

            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button("Nächste offene"):
                    if jump_next_open():
                        st.rerun()
                    else:
                        st.info("Alle Tätigkeiten haben bereits Stunden ✅")

            with c2:
                st.caption("🔵 FS  |  ⚪️ AP  |  🔸 Tätigkeit")

            # Tree als Expanders (ruhig & übersichtlich)
            for i_fs, fs in enumerate(tree):
                # Filter check
                fs_match = (search.strip() == "") or (search.lower() in fs["name"].lower())

                # Vorab: wenn Search aktiv ist, expandieren wir nur, wenn irgendwas darunter matcht
                def subtree_matches():
                    if fs_match:
                        return True
                    for ap in fs.get("aps", []):
                        if search.lower() in ap["name"].lower():
                            return True
                        for t in ap.get("tasks", []):
                            if search.lower() in t["name"].lower():
                                return True
                    return False

                if search.strip() and not subtree_matches():
                    continue

                # FS-Status: ✅ wenn alle Tätigkeiten darunter Stunden haben (wenn es überhaupt Tätigkeiten gibt)
                fs_task_keys = []
                for ap in fs.get("aps", []):
                    for t in ap.get("tasks", []):
                        fs_task_keys.append(task_key(fs, ap, t))
                if fs_task_keys and all(sum_hours_for_key(k) > 0 for k in fs_task_keys):
                    fs_status = "✅"
                else:
                    fs_status = "⚠️"

                expanded = (selected["fs"]["id"] == fs["id"]) if selected else False
                with st.expander(f"{fs_status} 🔵 {fs['name']}", expanded=expanded):
                    for i_ap, ap in enumerate(fs.get("aps", [])):
                        ap_match = (search.strip() == "") or (search.lower() in ap["name"].lower())
                        # AP-Header
                        st.markdown(f"**⚪️ {ap['name']}**")

                        # Tätigkeiten
                        for i_t, t in enumerate(ap.get("tasks", [])):
                            t_match = (search.strip() == "") or (search.lower() in t["name"].lower())
                            if search.strip() and not (fs_match or ap_match or t_match):
                                continue

                            k = task_key(fs, ap, t)
                            s = sum_hours_for_key(k)
                            status = "✅" if s > 0 else "⚠️"
                            suffix = f" ({round(s, 2)} h)" if s > 0 else ""
                            is_current = (selected["key"] == k)

                            label = f"{status} 🔸 {t['name']}{suffix}"
                            if is_current:
                                label = "➡️ " + label

                            if st.button(label, key=f"pick_{k}"):
                                # Auswahl setzen
                                st.session_state.hours_sel = {"fs": i_fs, "ap": i_ap, "t": i_t}
                                st.rerun()

        with right:
            st.subheader("Stunden pro Mitarbeiter")

            fs = selected["fs"]
            ap = selected["ap"]
            t = selected["t"]
            k = selected["key"]

            st.caption(f"FS: {fs['name']}  |  AP: {ap['name']}")
            st.markdown(f"### 🔸 {t['name']}")

            # Dataframe für Editor vorbereiten
            current_map = st.session_state.hours_by_task.get(k, {})
            df = pd.DataFrame({
                "Mitarbeiter": employees,
                "Stunden": [float(current_map.get(emp, 0.0) or 0.0) for emp in employees]
            })

            edited = st.data_editor(
                df,
                hide_index=True,
                use_container_width=True,
                disabled=["Mitarbeiter"],
                column_config={
                    "Stunden": st.column_config.NumberColumn(
                        "Stunden",
                        min_value=0.0,
                        step=0.5,
                        format="%.2f",
                        help="Stunden für diese Tätigkeit. 0 bedeutet keine Zuordnung."
                    )
                },
                key=f"hours_editor_{k}",
            )

            # Zusammenfassung
            total_hours = float(edited["Stunden"].fillna(0).sum())
            contributors = int((edited["Stunden"].fillna(0) > 0).sum())
            st.info(f"Summe: {round(total_hours, 2)} h  |  Mitarbeiter mit Stunden: {contributors}")

            # Actions
            cA, cB, cC = st.columns([1, 1, 2])
            with cA:
                save = st.button("Speichern")
            with cB:
                save_next = st.button("Speichern & nächste offene")

            def commit_hours():
                m = {}
                for _, row in edited.iterrows():
                    emp = row["Mitarbeiter"]
                    val = row["Stunden"]
                    try:
                        v = float(val)
                    except Exception:
                        v = 0.0
                    if v < 0:
                        v = 0.0
                    # wir speichern auch 0er (macht später UI stabiler)
                    m[emp] = v
                st.session_state.hours_by_task[k] = m

            if save:
                commit_hours()
                st.success("Gespeichert ✅")
                st.rerun()

            if save_next:
                commit_hours()
                # zur nächsten offenen springen
                cur_idx = 0
                for idx, it in enumerate(tasks):
                    if it["key"] == k:
                        cur_idx = idx
                        break
                order = list(range(cur_idx + 1, len(tasks))) + list(range(0, cur_idx + 1))
                jumped = False
                for idx in order:
                    it = tasks[idx]
                    if sum_hours_for_key(it["key"]) <= 0:
                        st.session_state.hours_sel = {"fs": it["i_fs"], "ap": it["i_ap"], "t": it["i_t"]}
                        jumped = True
                        break
                if not jumped:
                    st.success("Alles erledigt ✅")
                st.rerun()

        st.divider()

        # Weiter erst, wenn alle Tätigkeiten Stunden haben
        all_done = done == total

        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev)
        with col2:
            st.button("Weiter →", on_click=go_next, disabled=not all_done)

        if not all_done:
            st.caption("Hinweis: Bitte jeder Tätigkeit Stunden zuordnen (mind. ein Mitarbeiter > 0).")
#########################################################################################
    elif step == 6:
        st.header("Export")

        employees = st.session_state.get("employees", [])
        project_tree = st.session_state.get("project_tree", [])
        hours_by_task = st.session_state.get("hours_by_task", {})
        absences_by_employee = st.session_state.get("absences_by_employee", {})
        wj_start = st.session_state.get("wirtschaft_start", None)
        wj_end = st.session_state.get("wirtschaft_end", None)

        # Guards
        if not employees or wj_start is None or wj_end is None:
            st.warning("Bitte zuerst Wirtschaftsjahr und Mitarbeiter erfassen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        if not project_tree:
            st.warning("Bitte zuerst Projekte anlegen.")
            st.button("← Zurück", on_click=go_prev)
            st.stop()

        # --- Weniger ist mehr: 2 Einstellungen ---
        mode_label = st.selectbox(
            "Verteilung",
            ["Von Anfang nach Ende", "Von Ende nach Anfang", "Flach über das Jahr"],
        )

        style_label = st.selectbox(
            "Buchungsstil",
            ["Kleine Buchungen (gleichmäßiger)", "Große Buchungen (mehr pro Tag)"],
        )

        mode = {
            "Von Anfang nach Ende": "forward",
            "Von Ende nach Anfang": "backward",
            "Flach über das Jahr": "flat",
        }[mode_label]

        # nur 2 Presets (kein extra UI)
        if style_label.startswith("Kleine"):
            min_h, max_h = 1.0, 3.0
        else:
            min_h, max_h = 6.0, 8.0

        # Export bytes speichern
        if "export_xlsx_bytes" not in st.session_state:
            st.session_state.export_xlsx_bytes = None
            st.session_state.export_filename = None

        if st.button("Excel generieren"):
            # 1) tasks_per_employee + Legende
            tasks_per_employee, legend_rows = build_tasks_per_employee(project_tree, hours_by_task)

            # 2) Basis-rows (AT Feiertage fix)
            # -> build_base_rows sollte dafür holidays.country_holidays("AT", years=years) verwenden (ohne subdiv UI)
            rows = build_base_rows(employees, wj_start, wj_end, absences_by_employee)

            # 3) Stunden verteilen: neue Logik statt buchungsversuche
            active_tasks = prepare_active_tasks(tasks_per_employee)
            assign_hours(rows, active_tasks, mode=mode, min_hours=min_h, max_hours=max_h, earliest_hour=9, latest_hour=13)

            remaining_total = sum(
                t["remaining_hours"]
                for tasks in active_tasks.values()
                for t in tasks
            )

            # 4) In Excel schreiben (3 Sheets)
            df_main = pd.DataFrame(rows[1:], columns=rows[0])
            df_legend = pd.DataFrame(legend_rows) if legend_rows else pd.DataFrame(columns=["Code", "Forschungsschwerpunkt", "Arbeitspaket", "Tätigkeit"])

            # Summary: geplante Stunden vs. übrig
            summary_rows = []
            for emp in employees:
                planned = sum(t["hours"] for t in tasks_per_employee.get(emp, []))
                remaining = 0.0
                for t in active_tasks.get(emp, []):
                    remaining += float(t.get("remaining_hours", 0.0))
                summary_rows.append({"Mitarbeiter": emp, "Geplant (h)": round(planned, 2), "Nicht verteilt (h)": round(remaining, 2)})
            df_summary = pd.DataFrame(summary_rows)

            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df_main.to_excel(writer, index=False, sheet_name="Zeiterfassung")
                df_legend.to_excel(writer, index=False, sheet_name="Legende")
                df_summary.to_excel(writer, index=False, sheet_name="Summary")

            st.session_state.export_xlsx_bytes = bio.getvalue()

            year_tag = f"{wj_start.year}-{wj_end.year}" if wj_start.year != wj_end.year else f"{wj_start.year}"
            st.session_state.export_filename = f"zeiterfassung_{year_tag}.xlsx"

            if remaining_total > 0:
                st.warning(f"Nicht alle Stunden konnten verteilt werden. Rest: {round(remaining_total, 2)} h (Details im Summary-Sheet).")
            else:
                st.success("Excel erfolgreich generiert ✅")

        # Download anzeigen, sobald generiert
        if st.session_state.export_xlsx_bytes:
            st.download_button(
                "Excel herunterladen",
                data=st.session_state.export_xlsx_bytes,
                file_name=st.session_state.export_filename or "zeiterfassung.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()

        col1, col2 = st.columns(2)
        with col1:
            st.button("← Zurück", on_click=go_prev)
        with col2:
            st.button("Neu starten", on_click=lambda: st.session_state.update({"step": 0}))
#########################################################################################

if __name__ == "__main__":
    main()
