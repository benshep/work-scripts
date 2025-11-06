import os

user_profile = os.path.expanduser('~')
music_folder = os.path.join(user_profile, 'Music')
docs_folder = os.path.join(user_profile, 'STFC', 'Documents')
downloads_folder = os.path.join(user_profile, 'Downloads')
misc_folder = os.path.join(user_profile, 'Misc')
sharepoint_folder = os.path.join(user_profile, 'Science and Technology Facilities Council')
hr_info_folder = os.path.join(user_profile, 'UKRI', 'Science and Technology Facilities Council - HR')
budget_folder = os.path.join(docs_folder, 'Budget')
