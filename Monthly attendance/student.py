import streamlit as st
import pandas as pd
import tempfile

# Function to generate roll numbers based on prefix and total count
def generate_roll_numbers(prefix, total_students):
    return [f"{prefix}{str(i + 1).zfill(2)}" for i in range(total_students)]

def main():
    st.title("Student Details Input")

    roll_prefix = st.text_input("Enter Roll Number Prefix (e.g., 22ECA):")
    has_discontinuities = st.selectbox("Are there any discontinuities in the roll numbers?", ["Yes", "No"])
    student_names = []
    valid_roll_numbers = []

    # Always ask for the total number of students and their names
    total_students = st.number_input("Enter Total Number of Students:", min_value=1)
    student_input = st.text_area("Enter student names (one per line):")

    # If there are discontinuities, ask for the discontinuous roll numbers
    if has_discontinuities == "Yes":
        discontinuous_rolls = st.text_area("Enter the discontinuous roll numbers (one per line):")
    else:
        discontinuous_rolls = ""

    if st.button("Submit"):
        student_names = [name.strip() for name in student_input.splitlines() if name.strip()]
        discontinuous_rolls = [f"{roll_prefix}{roll.strip().zfill(2)}" for roll in discontinuous_rolls.splitlines() if roll.strip()]
        valid_roll_numbers = generate_roll_numbers(roll_prefix, total_students)
        
        # Filter out discontinuous roll numbers if any
        if has_discontinuities == "Yes":
            valid_roll_numbers = [roll for roll in valid_roll_numbers if roll not in discontinuous_rolls]

        if len(student_names) <= len(valid_roll_numbers):
            # Create a DataFrame with roll numbers and names
            df = pd.DataFrame({"Roll No": valid_roll_numbers[:len(student_names)], "Name": student_names})
            df.insert(0, "S.No", range(1, len(df) + 1))
            st.write("Student Details:")
            st.dataframe(df)

            # Save the DataFrame to an Excel file
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
                file_path = temp_file.name
                df.to_excel(file_path, index=False, sheet_name="Students")

            # Provide a download button for the Excel file
            with open(file_path, "rb") as file:
                st.download_button(label="Download Excel", data=file, file_name="student_details.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("The number of student names must not exceed the available roll numbers after considering discontinuities.")

if __name__ == "__main__":
    main()
