from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx2pdf import convert

import forms_retrieval
import send_email

#chama a função para obter resposta do formulário de declaracoes de ausencias
new_dict=forms_retrieval.get_answers("1KzEWtwWcn7CgI2Sx40JZHApZ4-HqxOXQLVBYDS540gc")

#numero da variavel
num=0

def get_variables(num):
  #nome
  #cc
  # motivo
  # Inicio ausencia
  # Fim ausencia

  nome=list(list(new_dict["responses"][num]["answers"].values())[0]["textAnswers"]["answers"][0].values())[0]
  cc=list(list(new_dict["responses"][num]["answers"].values())[1]["textAnswers"]["answers"][0].values())[0]
  data_inicio=list(list(new_dict["responses"][num]["answers"].values())[2]["textAnswers"]["answers"][0].values())[0]
  motivo=list(list(new_dict["responses"][num]["answers"].values())[3]["textAnswers"]["answers"][0].values())[0]
  data_fim=list(list(new_dict["responses"][num]["answers"].values())[4]["textAnswers"]["answers"][0].values())[0]
  results=[nome,cc,data_inicio, motivo,data_fim]
  return (results)


nome=get_variables(num)[0]
print("nome",nome)
cc=get_variables(num)[1]
print("cc",cc)
data_inicio=get_variables(num)[2]
print("data_inicio",data_inicio)
motivo=get_variables(num)[3]
print("motivo",motivo)
data_fim=get_variables(num)[4]
print("datafim",data_fim)


#escrita do documento

document = Document(('/Users/danielmdias/docs/_fun_time/clinical/clinical_py/template.docx'))

#styles
style1 = document.styles['Normal']
font = style1.font
font.name = 'Calibri'
font.size = Pt(14)


r  = "Eu, Daniel Martinho Dias, médico, portador da Cédula Profissional nº 63783, venho por este meio declarar que {0}, CC/BI {1}, se encontrou impossibilitada de exercer as suas atividades profissionais desde {2} até {3} por motivos de doença.\nPor ser verdade e me ter sido pedido, emito a presente declaração que dato e assino.".format(nome, cc, data_inicio,data_fim)

p = document.add_paragraph(r)
p.paragraph_format.space_before = Pt(15)
p.paragraph_format.space_after = Pt(15)
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.style = style1

path='/Users/danielmdias/docs/_fun_time/clinical/clinical_py/declaracao_ausencia.docx'
document.save(path)


convert(path)

send_email.gmail_send_message_with_attachment("assina a requisição", '/Users/danielmdias/docs/_fun_time/clinical/clinical_py/declaracao_ausencia.pdf')