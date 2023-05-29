from docx import Document
from html import escape

def convert_docx_to_html(docx_file):
    doc = Document(docx_file)
    html_content = "<html><body>"

    for para in doc.paragraphs:
        style = para.style.name

        if para.runs:
            # Check if the paragraph has bullet points
            if para.runs[0].font.bold and para.runs[0].text == "•":
                html_content += "<ul>"
                bullet_level = 0

                # Iterate over runs in the paragraph
                for run in para.runs:
                    run_text = run.text.strip()

                    if run.font.bold and run.text == "•":
                        bullet_level += 1
                        html_content += "<ul>"
                    elif bullet_level > 0:
                        if run_text:
                            html_content += f"<li>{escape(run_text)}</li>"
                        if run.font.bold and run.text == "•":
                            bullet_level -= 1
                            html_content += "</ul>"

                html_content += "</ul>"
            else:
                # Normal paragraph text
                text = para.text
                escaped_text = escape(text)

                if style == "Heading 1":
                    html_content += f"<h1>{escaped_text}</h1>"
                elif style == "Heading 2":
                    html_content += f"<h2>{escaped_text}</h2>"
                elif style == "Heading 3":
                    html_content += f"<h3>{escaped_text}</h3>"
                else:
                    html_content += f"<p>{escaped_text}</p>"

    html_content += "</body></html>"
    return html_content

docx_file_path = "input.docx"
html_content = convert_docx_to_html(docx_file_path)
html_file_path = "output.html"
with open(html_file_path, "w", encoding="utf-8") as file:
    file.write(html_content)

print(f"Conversion complete. HTML file saved at {html_file_path}.")
