import streamlit as st
import pandas as pd
from io import BytesIO


def find_chain(df, start_index):
    chain = []
    visited = set()
    current_index = start_index

    while current_index not in visited and current_index < len(df):
        visited.add(current_index)
        current_person = df.iloc[current_index]
        current_location = current_person["Görev Yaptığı Yer"]

        for preference in ["Tercih 1", "Tercih 2", "Tercih 3", "Tercih 4", "Tercih 5"]:
            preferred_location = current_person[preference]
            if pd.isna(preferred_location):
                continue

            next_person = df[df["Görev Yaptığı Yer"] == preferred_location]
            if not next_person.empty:
                next_index = next_person.index[0]
                if preferred_location == df.iloc[start_index]["Görev Yaptığı Yer"]:
                    chain.append((current_index, next_index))
                    return chain
                chain.append((current_index, next_index))
                current_index = next_index
                break
        else:
            return []

    return chain


def process_chain(df, chain):
    if not chain:
        return df, [], {}

    new_assignments = []
    movement_counts = {}

    for current_index, next_index in chain:
        current_person = df.iloc[current_index]
        next_person = df.iloc[next_index]

        old_location = current_person["Görev Yaptığı Yer"]
        new_location = next_person["Görev Yaptığı Yer"]

        if old_location not in movement_counts:
            movement_counts[old_location] = {"Gelen": 0, "Giden": 0}
        if new_location not in movement_counts:
            movement_counts[new_location] = {"Gelen": 0, "Giden": 0}

        movement_counts[old_location]["Giden"] += 1
        movement_counts[new_location]["Gelen"] += 1

        new_assignments.append(
            [
                current_person["Sicil"],
                current_person["Adı Soyadı"],
                current_person["Görev Yaptığı Yer"],
                new_location,
            ]
        )

    drop_indices = [current_index for current_index, _ in chain]
    df.drop(df.index[drop_indices], inplace=True)
    df.reset_index(drop=True, inplace=True)

    return df, new_assignments, movement_counts


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df.to_excel(writer, index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


st.title("Sonsuz Becayiş Uygulaması")

uploaded_file = st.file_uploader("Excel dosyasını yükleyin", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    original_df = df.copy()

    successful_assignments = []
    overall_movement_counts = {}

    for i in range(len(df)):
        if i >= len(df):
            break
        chain = find_chain(df, i)
        df, new_assignments, movement_counts = process_chain(df, chain)
        if new_assignments:
            successful_assignments.extend(new_assignments)

            for location, counts in movement_counts.items():
                if location not in overall_movement_counts:
                    overall_movement_counts[location] = {"Gelen": 0, "Giden": 0}
                overall_movement_counts[location]["Gelen"] += counts["Gelen"]
                overall_movement_counts[location]["Giden"] += counts["Giden"]

    if successful_assignments:
        successful_df = pd.DataFrame(
            successful_assignments,
            columns=["Sicil", "Adı Soyadı", "Eski Görev Yeri", "Yeni Görev Yeri"],
        )
    else:
        successful_df = pd.DataFrame(
            columns=["Sicil", "Adı Soyadı", "Eski Görev Yeri", "Yeni Görev Yeri"]
        )

    unsuccessful_df = original_df[~original_df["Sicil"].isin(successful_df["Sicil"])]

    if overall_movement_counts:
        movement_counts_df = pd.DataFrame.from_dict(
            overall_movement_counts, orient="index"
        ).reset_index()
        movement_counts_df.columns = ["Görev Yeri", "Gelen", "Giden"]
    else:
        movement_counts_df = pd.DataFrame(columns=["Görev Yeri", "Gelen", "Giden"])

    st.write("Başarılı Atamalar")
    st.write(successful_df)

    st.write("Atanamayanlar")
    st.write(unsuccessful_df)

    st.write("Görev Yerleri Hareketleri")
    st.write(movement_counts_df)

    successful_excel = to_excel(successful_df)
    unsuccessful_excel = to_excel(unsuccessful_df)
    movement_counts_excel = to_excel(movement_counts_df)

    st.download_button(
        "Başarılı Atamaları İndir",
        data=successful_excel,
        file_name="basarili_atamalar.xlsx",
    )
    st.download_button(
        "Atanamayanları İndir", data=unsuccessful_excel, file_name="atanamayanlar.xlsx"
    )
    st.download_button(
        "Görev Yerleri Hareketlerini İndir",
        data=movement_counts_excel,
        file_name="hareket_raporu.xlsx",
    )
