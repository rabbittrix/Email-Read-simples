using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

/* A ideia do sistema é acessar os e-mails, identificar as palavras chaves, baixar os anexos, 
 * identificar palavras chaves nos anexos, classificar os e-mails e anexos */

namespace OneyEmailRead
{
    public partial class Form1 : Form
    {
        private List<Email> _emails = new List<Email>();
        private string _hostname = "outlook.office365.com"; // Host do seu servidor POP3. Por exemplo, pop.gmail.com para o servidor do Gmail.
        private int _port = 995; // Porta utilizada pelo host. Por exemplo, 995 para o servidor do Gmail.
        private bool _useSsl = true; // Indicação se o servidor POP3 utiliza SSL para autenticação. Por exemplo, o servidor do Gmail utiliza SSL, então, "true".
        private string _username = "testes@cc.oney.pt"; // Usuário do servidor POP3. Por exemplo, seuemail@gmail.com.
        private string _password = "Waz60711"; // Senha do servidor POP3.
        private string arquivo;
        private string mensagem;

        public Form1()
        {
            InitializeComponent();
        }

        /* classe servirá para representar cada e-mail que será baixado do servidor.
         * classe para utilizarmos no data-binding do ListBox e TextBoxes: */
        public class Email
        {
            public string Id { get; set; }
            public string Assunto { get; set; }
            public string De { get; set; }
            public string Para { get; set; }
            public DateTime Data { get; set; }
            public string ConteudoTexto { get; set; }
            public string ConteudoHtml { get; set; }
        }

        /* código do evento “Click” do botão “Carregar“ */
        protected void Read_Emails(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                using (var client = new OpenPop.Pop3.Pop3Client())
                {
                    client.Connect(_hostname, _port, _useSsl);
                    client.Authenticate(_username, _password);
                    int messageCount = client.GetMessageCount();
                    _emails.Clear();

                    for (int i = messageCount; i > 0; i--)
                    {
                        var popEmail = client.GetMessage(i);

                        /* Para recuperarmos a versão da mensagem em formato texto, 
                         * utilizamos o método “FindFirstPlainTextVersion“. 
                         * Já para baixarmos a versão HTML, utilizamos o método “FindFirstHtmlVersion“ */
                        var popText = popEmail.FindFirstPlainTextVersion();
                        var popHtml = popEmail.FindFirstHtmlVersion();

                        string mailText = string.Empty;
                        string mailHtml = string.Empty;
                        if (popText != null)
                            mailText = popText.GetBodyAsText();
                        if (popHtml != null)
                            mailHtml = popHtml.GetBodyAsText();

                        /*  agora que já temos a mensagem tanto em formato texto quanto em formato HTML,
                         *  podemos armazená-las dentro da nossa lista de e-mails 
                         *  (atributo chamado “_emails” no nível do formulário) */
                        _emails.Add(new Email()
                        {
                            Id = popEmail.Headers.MessageId,
                            Assunto = popEmail.Headers.Subject,
                            De = popEmail.Headers.From.Address,
                            Para = string.Join("; ", popEmail.Headers.To.Select(to => to.Address)),
                            Data = popEmail.Headers.DateSent,
                            ConteudoTexto = mailText,
                            ConteudoHtml = !string.IsNullOrWhiteSpace(mailHtml) ? mailHtml : mailText

                        });
                                               
                    }

                    AtualizarDataBindings();

                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        /* Implementa uma pesquisa no RichTextBox */
        public int FindText(string searchText, int searchStart, int searchEnd)
        {
            // Inicialize o valor de retorno para false por padrão.
            int returnValue = -1;

            // Assegure-se de que uma cadeia de pesquisa e um ponto de início válido sejam especificados.
            if (searchText.Length > 0 && searchStart >= 0)
            {
                // Assegure-se de que um valor final válido seja fornecido.
                if (searchEnd > searchStart || searchEnd == -1)
                {
                    // Obtenha a localização da cadeia de pesquisa em richTextBox.
                    int indexToText = conteudoTextBox.Find(searchText, searchStart, searchEnd, RichTextBoxFinds.MatchCase);
                    // Determine se o texto foi encontrado em richTextBox1.
                    if (indexToText >= 0)
                    {
                        // Retorna o Index para a pesquisa do texto especifico.
                        returnValue = indexToText;
                    }
                }
            }

            return returnValue;
        }
        private void AtualizarDataBindings()
        {
            // Limpando os bindings.
            deTextBox.DataBindings.Clear();
            paraTextBox.DataBindings.Clear();
            dataDateTimePicker.DataBindings.Clear();
            conteudoTextBox.DataBindings.Clear();
            emailsListBox.DataSource = null;
            emailsBindingSource.DataSource = null;

            // Re-configurando os bindings.
            emailsBindingSource.DataSource = _emails;
            emailsListBox.DataSource = emailsBindingSource;
            emailsListBox.DisplayMember = "Assunto";
            deTextBox.DataBindings.Add(new Binding("Text", emailsBindingSource, "De"));
            paraTextBox.DataBindings.Add(new Binding("Text", emailsBindingSource, "Para"));
            dataDateTimePicker.DataBindings.Add(new Binding("Value", emailsBindingSource, "Data"));
            conteudoTextBox.DataBindings.Add(new Binding("Text", emailsBindingSource, "ConteudoTexto"));
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
                /* Metodo para abrir e ler cada linha de um arquivo, isto vai servir para 
                    identificar algumas palavras chaves no arquivo.*/

          /*if (File.Exists(@"D:\Criar roteador notebook.txt"))
            {
                Stream entrada = File.Open(@"D:\Criar roteador notebook.txt", FileMode.Open);
                StreamReader leitor = new StreamReader(entrada);
                string linha = leitor.ReadLine();
                while (linha != null)
                {
                    MessageBox.Show(linha);
                    linha = leitor.ReadLine();
                }
                leitor.Close();
                entrada.Close();
            }*/

            // buscar uma palavra dentro de um arquivo de texto
            {
                String varPalavra = "Sublime";

                StreamReader re = File.OpenText(@"D:\Criar roteador notebook.txt");
                string input = re.ReadToEnd();

                if (input.IndexOf(varPalavra) > -1)
                    MessageBox.Show("Existe a palavra '" + varPalavra + "' no arquivo txt");
                    
                else
                    MessageBox.Show("Não existe a palavra '" + varPalavra + "' no arquivo txt");

                re.Close();

            }
            

            /* Metodo abre o arquivo escolhido e armazena em um array para,
               podendo verificar cada palavra, letra ou frase. 
            {
                    List<string> mensagemLinha = new List<string>();
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Title = "Lendo arquivo";
                        openFileDialog.InitialDirectory = @"D:\Criar roteador notebook.txt"; //Se ja quiser em abrir em um diretorio especifico
                        openFileDialog.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                        openFileDialog.FilterIndex = 2;
                        openFileDialog.RestoreDirectory = true;
                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                            arquivo = openFileDialog.FileName;
                    }
                    if (String.IsNullOrEmpty(arquivo))
                    {
                        MessageBox.Show("Arquivo Invalido", "Salvar Como", MessageBoxButtons.OK);
                    }
                    else
                    {
                        using (StreamReader texto = new StreamReader(arquivo))
                        {
                            while ((mensagem = texto.ReadLine()) != null)
                            {
                                mensagemLinha.Add(mensagem);
                            }
                        }
                        int registro = mensagemLinha.Count; //total de linhas do arquivo.
                        for (int i = 0; i < mensagemLinha.Count; i++)
                        {
                            TextBox textbox1 = new TextBox();
                            textbox1.Text += mensagemLinha[i];
                            File.WriteAllText(arquivo, mensagemLinha[i] + "1");
                        }
                                        
                    }
            } */
            
        }
    }

        /*
         * Este exemplo procura a string "office" em Body de itens na caixa de entrada.
         * using Outlook = Microsoft.Office.Interop.Outlook;
         * private void SearchBody()
            {
                string filter;
                if (Application.Session.DefaultStore.IsInstantSearchEnabled)
                {
                    filter = "@SQL=" + "\""
                        + "urn:schemas:httpmail:textdescription" + "\""
                        + " ci_phrasematch 'office'";
                }
                else
                {
                    filter = "@SQL=" + "\""
                        + "urn:schemas:httpmail:textdescription" + "\""
                        + " like '%office%'";
                }
                Outlook.Table table = Application.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox).GetTable(
                    filter, Outlook.OlTableContents.olUserItems);
                while (!table.EndOfTable)
                {
                    Outlook.Row row = table.GetNextRow();
                    Debug.WriteLine(row["Subject"]);
                }
            }

               ******************************************************
             " ******** Metodo para salvar anexos do Outlook ******** "
            using SWF = System.Windows.Forms;
            using Microsoft.Office.Interop.Outlook;
            using System.Text.RegularExpressions;


            namespace SalvaAnexosOutlook
            {
                public partial class Form1 : SWF.Form
                {
                    public Form1()
                    {
                        InitializeComponent();
                    }


                    private void Form1_Load(object sender, EventArgs e)
                    {
                        Application myApp = new ApplicationClass();


                        foreach (Folder f in myApp.GetNamespace("MAPI").Folders)
                            listBox1.Items.Add(f.Name);
                        listBox1.SelectedIndex = 0;
                    }

                    private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
                    {
                        if (listBox1.SelectedIndex != -1)
                        {
                            listBox2.Items.Clear();


                            Application myApp = new ApplicationClass();


                            foreach (Folder f in myApp.GetNamespace("MAPI")
                                .Folders[listBox1.SelectedIndex + 1].Folders)
                                listBox2.Items.Add(f.Name);
                        }
                    }


                    private void salvarAnexo(object sender, EventArgs e)
                    {
                        if (SWF.MessageBox.Show(
                            "Tem certeza que quer salvar os anexos das mensagens selecionadas?" +
                            "\n\nIsso apagará os anexos do seu e-mail!!!", "Responda",
                            SWF.MessageBoxButtons.YesNo,
                            SWF.MessageBoxIcon.Question)
                                == SWF.DialogResult.Yes)
                        {
                            Application myApp = new ApplicationClass();


                            foreach (Folder f in myApp.GetNamespace("MAPI")
                                .Folders[listBox1.SelectedIndex + 1].Folders)
                            {
                                if (f.Name.Equals(listBox2.SelectedItem.ToString()))
                                {
                                    foreach (object email in (f.Items))
                                    {
                                        MailItem mail = email as MailItem;
                                        
                                        int contador;
                                        
                                        if (mail != null)
                                        {
                                            contador = 0;
                                            
                                            while (mail.Attachments.Count > 0 && contador <= 10)
                                            {
                                                Attachment mi = mail.Attachments[1] as Attachment;

                                                if (mi != null)
                                                {
                                                    string nomeAtt = null;
                                                    
                                                    try
                                                    {
                                                        nomeAtt = SWF.Application.StartupPath
                                                            + @"\" + f.Name
                                                            + "_" + mail.SentOn.ToString("yyyyMMddHHmmss")
                                                            + "_" + Limpar(mail.Subject)
                                                            + "_" + Limpar(mi.DisplayName);
                                                        mi.SaveAsFile(nomeAtt);
                                                        
                                                        label1.Text = mi.DisplayName;
                                                        SWF.Application.DoEvents();
                                                        
                                                        mi.Delete();
                                                    }
                                                    catch
                                                    {
                                                        contador++;
                                                    }
                                                }
                                            }
                                            
                                            if (mail != null)
                                                mail.Save();
                                        }
                                    }

                                    if (SWF.MessageBox.Show(
                                        "Arquivos gravados - quer abrir a pasta?",
                                        "FIM DO PROCESSO",
                                        SWF.MessageBoxButtons.YesNo,
                                        SWF.MessageBoxIcon.Question)
                                            == SWF.DialogResult.Yes)
                                    {
                                        System.Diagnostics.Process.Start(
                                            SWF.Application.StartupPath);
                                        SWF.Application.Exit();
                                    }
                                }
                            }
                        }
                    }


                    private String Limpar(String s)
                    {
                        return Regex.Replace(s, "[\"@/{}<>():|;\\?*&%$]", String.Empty);
                    }

                    private void linkLabel1_LinkClicked(object sender, SWF.LinkLabelLinkClickedEventArgs e)
                    {
                        System.Diagnostics.Process.Start(roberto.souza@ctr.unipartner.com);
                    }
                }
            }
         */
    }

