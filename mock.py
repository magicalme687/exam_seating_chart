import pandas as pd

students = {
    'I Yr': [f'IS{i:03}' for i in range(1, 41)],
    'II Yr': [f'IIS{i:03}' for i in range(1, 41)],
    'III Yr': [f'IIIS{i:03}' for i in range(1, 26)],
    'IV Yr': [f'IVS{i:03}' for i in range(1, 21)]
}

df_students = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in students.items()]))
df_students.to_excel('students_test.xlsx', index=False)

subjects = {
    'I Yr': ['CS101', 'CS102', 'CS103'],
    'II Yr': ['CS201', 'CS202', 'CS203'],
    'III Yr': ['CS301', 'CS302', 'CS303'],
    'IV Yr': ['CS401', 'CS402', 'CS403']
}

df_subjects = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in subjects.items()]))
df_subjects.to_excel('subjects_test.xlsx', index=False)
