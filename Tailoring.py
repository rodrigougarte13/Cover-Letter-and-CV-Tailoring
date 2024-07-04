import openai
import pandas as pd
import docx
from docx import Document

openai.api_key = 'sk-proj-UJa1VcmmtY4NjEfOxutaT3BlbkFJRe1w6T9WddxOS6rQBDr0'


def extract_text(file_path):
    doc = Document(file_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return "\n".join(full_text)


# CV GENERATION
def generate_cv(cv_path, job_offer):
    cv_text = read_cv(cv_path)
    new_bullets_text = update_cv_sections(cv_text, job_offer)
    return update_cv(cv_path, new_bullets_text)


# COVER LETTER GENERATION
def generate_cover_letter(company, position, job_offer):
    cover_letter_path = 'Cover Letter.docx'
    # cover letter text
    cover_letter_template = extract_text(cover_letter_path)

    # prompt for GPT-4o
    prompt_text = (
        f"Using the cover letter template and Job Offer Details provided below, craft a complete, compelling cover letter. "
        f"The cover letter should articulate why the applicant is drawn to {company}, and fit to be a {position}and showcase my skills, experiences and competences "
        f"Leverage all my qualities, also showing my softskills, and dont invent false information "
        f"specifically citing relevant company values and initiatives mentioned in the job offer but mainly ones found online. "
        f"Keep it under 470 words so trim information on the template that is not relevant to the job offer and make sure you return a cover letter from a person you would definitely hire."
        f"Maintain a professional writing style and match the writing style of the template"
        f"Moreover, I am attacching bullet points from my cv, please adapt them with the info of the cover letter and my cover letter\n\n"
        f"Job Offer Details: {job_offer}\n\n"
        f"Cover Letter Template: {cover_letter_template}")

    # API call
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "system",
                   "content": f"You are a recruiter with 20 years of experience in big tech companies, expert in CV and cover letter taloring"},
                  {"role": "user", "content": prompt_text}],
        max_tokens=620)

    return response.choices[0].message.content


def tailor():
    job_df = pd.read_csv('Cover Letter List.csv', encoding='latin-1')
    for index, row in job_df.iterrows():
        company = row['Company']
        position = row['Position']
        job_offer = row['Job Offer']
        cover_letter = generate_cover_letter(company, position, job_offer)

        # save the generated cover letter to a file
        doc = Document()
        doc.add_paragraph(cover_letter)
        doc.save(f'Cover_Letter_{company}_{position}.docx')
        print(f'Cover Letter for {company} for {position} done!')

        # save CV to a file
        cv_path = 'CV Rodrigo Ugarte.docx'
        cv = generate_cv(cv_path, job_offer)
        cv.save(f'CV_Rodrigo_Ugarte_{company}_{position}.docx')
        print(f'CV for {company} for {position} done!')