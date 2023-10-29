################ ENVIAR EMAIL #################3

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using SAPbouiCOM.Framework;

namespace AGN_Email
{
    [FormAttribute("AGN_Email.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {

        private List<string> selectedValues = new List<string>();

        public SAPbobsCOM.Company oCompany;

        public Form1(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }

        // CHECKBOX
        public Form1()
        {
            Grid0.Columns.Item("Selecione").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
        }

        // SELECT 'N' AS "Selecione", T0."DocNum" AS "Nº Pedido", T0."CardName" AS "Cliente", T0."DocDate" AS "Data", T0."Email" AS "E-mail" FROM OPOR T0
        // SELECT 'N' AS "Selecione", T0."DocNum" AS "Nº Pedido", T0."CardName" AS "Cliente", T0."DocDate" AS "Data", T1."E_Mail" AS "E-mail" FROM OPOR T0 INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode"

        private SAPbouiCOM.Grid Grid1;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_1").Specific));
            this.Grid0.ComboSelectBefore += new SAPbouiCOM._IGridEvents_ComboSelectBeforeEventHandler(this.Grid0_ComboSelectBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            //   this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {


        }

        private void OnCustomInitialize()
        {

        }

        private void Grid1_DatasourceLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private SAPbouiCOM.Grid Grid0;
        private void Grid0_DatasourceLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            throw new System.NotImplementedException();

        }

        private void Grid0_ComboSelectBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            throw new System.NotImplementedException();
        }

        private SAPbouiCOM.Button Button1;


// BOTÃO
        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            List<string> destinatarios = new List<string>();

            for (int i = 0; i < Grid0.Rows.Count; i++)
            {
                string checkBoxValue = Grid0.DataTable.GetValue("Selecione", i) as string;

                //  "Y" = marcado
                if (checkBoxValue != null && checkBoxValue.Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    string NPedido = Grid0.DataTable.GetValue("Nº Pedido", i).ToString();
                    selectedValues.Add(NPedido);

                    string email = Grid0.DataTable.GetValue("E-mail", i).ToString();

                    destinatarios.Add(email);

                    Console.WriteLine("Nº de pedido selecionados: " + string.Join(", ", selectedValues));
                }
            }

            if (destinatarios.Count > 0)    
            {
                EnviarEmailGmail(destinatarios, "Testando!", "Isto é um teste de email para os Nº de Pedidos selecionados.");
            }           
        }

        private void EnviarEmailGmail(List<string> destinatarios, string assunto, string corpo)
        {
            try
            {
                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com"))
                {
                    smtpClient.Port = 587;
                    smtpClient.Credentials = new NetworkCredential("t1769320@gmail.com", "etzl cljv gypr fggx");
                    smtpClient.EnableSsl = true;

                    using (MailMessage message = new MailMessage())
                    {
                        message.From = new MailAddress("t1769320@gmail.com");

                        foreach (string destinatario in destinatarios)
                        {
                            message.To.Add(destinatario);
                        }

                        message.Subject = assunto;
                        message.Body = corpo;
                        message.IsBodyHtml = true;

                        smtpClient.Send(message);

                        MessageBox.Show("Emails enviados para os destinatários selecionados!");
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Erro ao enviar o email:  " + ex.Message);

        }
      }
   }
}
