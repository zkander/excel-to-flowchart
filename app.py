import streamlit as st
import pandas as pd
import subprocess

def parse_excel(file):
    df = pd.read_excel(file)
    return df

def generate_mermaid(df):
    mermaid_code = "graph TD\n"
    for index, row in df.iterrows():
        process_step_id = row['Process Step ID']
        process_step_description = row['Process Step Description']
        next_step_ids = str(row['Next Step ID']).split(',')
        connector_labels = str(row['Connector Label']).split(',')
        
        if row['Shape Type'] == 'Decision':
            mermaid_code += f"{process_step_id}[\"{process_step_description}\"]\n"
        else:
            mermaid_code += f"{process_step_id}[{process_step_description}]\n"
        
        for i, next_step_id in enumerate(next_step_ids):
            if next_step_id and next_step_id != 'nan':
                if len(connector_labels) > i and connector_labels[i] and connector_labels[i] != 'nan':
                    mermaid_code += f"{process_step_id} --> |{connector_labels[i]}| {next_step_id}\n"
                else:
                    mermaid_code += f"{process_step_id} --> {next_step_id}\n"

    return mermaid_code

def main():
    st.title("Excel to Flowchart")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        df = parse_excel(uploaded_file)
        st.dataframe(df)

        mermaid_code = generate_mermaid(df)
        #Add the mermaid code to an md file
        file = open("output.mmd", "w")
        file.write(mermaid_code)
        file.close()
        command = "mmdc -i output.mmd -o output.svg -t dark -b transparent" 
        subprocess.run(command, shell=True)
        
        with open("output.svg", "r") as file:
            svg = file.read()

        # Add a style attribute to set width to 100%
        svg_with_width = svg.replace('<svg', '<svg style="width:50%;" ')

        centered_svg = f'<div style="display: flex; justify-content: center;">{svg_with_width}</div>'

        st.markdown(centered_svg, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
