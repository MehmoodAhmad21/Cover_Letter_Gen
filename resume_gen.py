from docx import Document

def generate_cover_letter(company_name, your_name, Degree, School, Extraciriculars, Program_Languages, frameworks, dev_tools, libraries, Previous_Role, What_you_did,):
    
    letter_template = f"""
Dear Hiring Manager,
I am writing to express my keen interest in the position at {company_name}, as advertised on your company's website. Currently pursuing a {Degree} the {School}, I am enthusiastic about the prospect of contributing my diverse skill set and innovation-driven mindset to {company_name}'s team.\n
Alongside my academic pursuits, I actively {Extraciriculars}. These experiences have honed my organizational and teamwork skills, making me well-equipped to thrive in a collaborative work environment.\n
My technical proficiency spans a range of programming languages, including {Program_Languages}. I have hands-on experience with frameworks like {frameworks}, leveraging developer tools such as {dev_tools}. Additionally, I am adept in various libraries, including {libraries}.\n
In my previous role as a {Previous_Role}, I successfully {What_you_did}. This experience, coupled with strong problem-solving skills, allowed me to excel in a dynamic and collaborative environment.
The prospect of contributing my technical expertise to {company_name}, a company known for its commitment to excellence and innovation in natural resources, is truly exciting. I am drawn to your organization's forward-thinking approach and commitment to pushing boundaries, aligning seamlessly with my own dedication to staying up to date with the latest technologies and methodologies.\n
Confident in my technical skills, collaborative mindset, and commitment to continuous learning, I believe I would be a valuable addition to your team. I am eager to contribute to {company_name}'s initiatives and help drive the company's success.
Thank you for considering my application. I look forward to the opportunity to discuss how my skills and experiences align with the needs of your team.\n
Sincerely,
{your_name}
"""

    return letter_template

def create_word_document(cover_letter_text, output_file="CoverLetter.docx"):
    document = Document()
    document.add_paragraph(cover_letter_text)

    document.save(output_file)
    print(f"Cover letter saved to {output_file}")

# Example usage:
company_name = input("Enter the company name: ")
# All relevent data, Write here 
your_name = "John Doe" # Replace String with data relevent to you
Degree = "Degree Name" 
School = "School Name"
Extraciriculars = "Extraciriculars"
Program_Languages = "Python, Html, ....."
frameworks = "Django, ...., etc"
dev_tools = "Vscode, .., etc."
libraries = "Pandas, ..... ,etc"
Previous_Role = "Software Engineer"
What_you_did = "I sucefully completed this and that and so one"




cover_letter = generate_cover_letter(company_name, your_name, Degree, School, Extraciriculars, Program_Languages, frameworks, dev_tools, libraries, Previous_Role, What_you_did)
print(cover_letter)
create_word_document(cover_letter)

