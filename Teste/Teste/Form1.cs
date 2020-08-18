using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Teste
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

        //Método de busca e substituição
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Criando o Documento Word
        private void CreateWordDocument(object filename, object SaveAs, string nome)
        {
            
                
                Word.Application wordApp = new Word.Application();
                object missing = Missing.Value;
                Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();


                    //Buscando "<nome>" e substituindo pelo valor da variável "nome"
                    this.FindAndReplace(wordApp, "<nome>", nome);
                    
                }//CONTINUA ABAIXO...
                else
                {
                    MessageBox.Show("Arquivo não encontrado!");
                }

                //Salvando documento
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                
            
        }
        //Botão "Gerar"
        private void button1_Click(object sender, EventArgs e)
		{
            //pegamos a grande string que é escrita na textbox e separamos de acordo com as vírgulas da string
            string[] participantes = textBox1.Text.Split(new char[] { ',' });
            //contador comum
            int i;
          

            for (i = 0; i < participantes.Length; i++) { 
            //primeiro parâmetro: local do documento modelo junto do nome
            //segundo parâmetro: local onde deve ser salvo junto do nome
            //terceiro parâmetro: Nome que será colocado no certificado
            CreateWordDocument(@"C:\Users\samuq\source\repos\Teste\Teste\certificado.docx",
                @"C:\Users\samuq\Desktop\CertificadosPETRedação\certificado"+i+".docx", participantes[i]);
            }

            if(i == participantes.Length)
            {
                MessageBox.Show("Certificando gerados!");
            }
        }
	}
}
