using ModelodeAprov.Controller;
using ModelodeAprov.models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using DIAPI = SAPbobsCOM;

namespace ModelodeAprov
{
    public partial class Form1 : Form
    {

        

        private string status = "and  dd.Status in ('W')";
        private string squery;
 
        public Form1()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e) {

            //checkBox3.Checked = true;
            //AtualizarGrid(LoginUser.user);
            
             carregaGrid(LoginUser.user);

        }

        private void carregaGrid( string user)
        {

            if (ConectaSAP.oCompany.Connected)
            {
                DIAPI.Recordset oRecSetBuscarAprov = ConectaSAP.oCompany.GetBusinessObject(DIAPI.BoObjectTypes.BoRecordset);

                Pesquisar(LoginUser.user);

                try
                {
                    oRecSetBuscarAprov.DoQuery(squery);

                }
                catch (Exception e)
                {
                    string erro = e.Message;
                    MessageBox.Show(e.Message);
                }
                
                System.Data.DataTable dt = new System.Data.DataTable();

                //criar colunas no grid

                dt.Columns.Add("Status", typeof(string));
                dt.Columns.Add("Total", typeof(string));
                dt.Columns.Add("Código", typeof(string));
                dt.Columns.Add("Modelo", typeof(string));
                dt.Columns.Add("Usuario", typeof(string));
                dt.Columns.Add("Observação", typeof(string));
                dt.Columns.Add("PN", typeof(string));
                dt.Columns.Add("Anexo", typeof(string));


                var col = new DataGridViewCheckBoxColumn();
                col.Name = "Aprovar";
                col.HeaderText = "Aprovar";
                col.FalseValue = "0";
                col.TrueValue = "1";

                //Make the default checked
                col.CellTemplate.Value = false;
                col.CellTemplate.Style.NullValue = false;

                dgvDados.Columns.Insert(0, col);


                for (int i=0 ; i < oRecSetBuscarAprov.RecordCount; i++)
                {
                    var DadosAProv = new DadosAProv()
                    {
                        Codigo = oRecSetBuscarAprov.Fields.Item("Codigo").Value.ToString(),
                        Modelo = oRecSetBuscarAprov.Fields.Item("Modelo").Value.ToString(),
                        Usuario = oRecSetBuscarAprov.Fields.Item("Usuario").Value.ToString(),
                        Obervacao = oRecSetBuscarAprov.Fields.Item("Observacao").Value.ToString(),
                        PN = oRecSetBuscarAprov.Fields.Item("PN").Value.ToString(),
                        Total = oRecSetBuscarAprov.Fields.Item("Total").Value.ToString(),
                        Status = oRecSetBuscarAprov.Fields.Item("Status").Value.ToString(),
                        Anexo = oRecSetBuscarAprov.Fields.Item("Anexo").Value.ToString(),

                    };


                    //Preencher o grid
                    dt.Rows.Add(DadosAProv.Status, DadosAProv.Total,DadosAProv.Codigo,DadosAProv.Modelo, DadosAProv.Usuario,DadosAProv.Obervacao,DadosAProv.PN, DadosAProv.Anexo);
              
                    dgvDados.DataSource = dt;
                    oRecSetBuscarAprov.MoveNext();

                }

            }
          
        }

        private void button1_Click(object sender, EventArgs e)
        {

            int contalinhas = 0;
            foreach (DataGridViewRow row in dgvDados.Rows)
            {
                if (row.IsNewRow) continue;


                string selAprov = row.Cells["Status"].Value.ToString();
                int Code = Convert.ToInt32(row.Cells["Código"].Value.ToString());


                if (Convert.ToBoolean(row.Cells["Aprovar"].FormattedValue))
                {
                  

                    if (selAprov == "Pendente")
                    {


                        DIAPI.CompanyService cs = ConectaSAP.oCompany.GetCompanyService();
                        DIAPI.ApprovalRequestsService approvalSrv = cs.GetBusinessService(DIAPI.ServiceTypes.ApprovalRequestsService);
                        DIAPI.ApprovalRequestParams oParams = approvalSrv.GetDataInterface(DIAPI.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

                        approvalSrv.GetApprovalRequestList();

                        oParams.Code = Code;
                        DIAPI.ApprovalRequest oData = approvalSrv.GetApprovalRequest(oParams);

                        oData.ApprovalRequestDecisions.Add();
                        oData.ApprovalRequestDecisions.Item(0).ApproverUserName = LoginUser.user; //  "manager";
                        oData.ApprovalRequestDecisions.Item(0).ApproverPassword = LoginUser.password ;// "Evo@09";
                        oData.ApprovalRequestDecisions.Item(0).Status = DIAPI.BoApprovalRequestDecisionEnum.ardApproved;
                        oData.ApprovalRequestDecisions.Item(0).Remarks = "Aprovador por " + LoginUser.user;

                        try
                        {
                            approvalSrv.UpdateRequest(oData);
                            AtualizarGrid(LoginUser.user);

                        }
                        catch (Exception ex)
                        {
                            string errorexception = ex.Message;
                            MessageBox.Show(ex.Message, "Erro na aprovação", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        
                    }
                   
                }
                label1.Visible = true;
                label1.BackColor = System.Drawing.Color.Green;
                label1.Text = "Modelos Aprovados Com Sucesso!";
                contalinhas = dgvDados.Rows.Count + 1;

            }


        

            if ( contalinhas == 0)
            {
                MessageBox.Show("Selecionar um linha");
            }
 

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in dgvDados.Rows)
            //{
            //    row.Cells["Aprovar"].Value = 1;

            //}
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                foreach (DataGridViewRow row in dgvDados.Rows)
                {
                    row.Cells["Aprovar"].Value = 1;

                }

            }
            else
            {
                foreach (DataGridViewRow row in dgvDados.Rows)
                {
                    row.Cells["Aprovar"].Value = 0;

                }

            }
        }

        private void dgvDados_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            string pathtofile =dgvDados.CurrentCell.Value.ToString();


            try
            {
                System.Diagnostics.Process.Start(pathtofile);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                   
            }
           
       
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //Aprovado
            if (checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked)
            {
                status = "and  dd.Status in ('Y')";

            }
            //Pendente
            if (checkBox3.Checked && !checkBox2.Checked && !checkBox4.Checked)
            {
                status = "and  dd.Status in ('W')";

            }
            //Aprovado e pendente
            if (checkBox2.Checked && checkBox3.Checked && !checkBox4.Checked)
            {
                status = "and  dd.Status in ('Y','W')";

            }
            //todos
            if (!checkBox3.Checked && !checkBox2.Checked && checkBox4.Checked)
            {
                status = "and  dd.Status in ('Y','W','N','P','A','C')";

            }

            //aprovado, pendente e todos
            if (checkBox3.Checked && checkBox2.Checked && checkBox4.Checked)
            {
                status = "";

            }

            AtualizarGrid(LoginUser.user);
        }

        private void AtualizarGrid( string user)
        {

            if (ConectaSAP.oCompany.Connected)
            {
 
                Pesquisar(LoginUser.user);
                DIAPI.Recordset oRecSetBuscarAprov = ConectaSAP.oCompany.GetBusinessObject(DIAPI.BoObjectTypes.BoRecordset);

                try
                {
                    oRecSetBuscarAprov.DoQuery(squery);

                }
                catch (Exception e)
                {
                    string erro = e.Message;
                    MessageBox.Show(e.Message);
                }


                if (oRecSetBuscarAprov.RecordCount == 0)
                {
                    MessageBox.Show("Não Foram Encontrados Dados com os Critérios de Seleção","Atenção",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Rows.Clear();
                //criar colunas no grid
                dt.Columns.Add("Status", typeof(string));
                dt.Columns.Add("Total", typeof(string));
                dt.Columns.Add("Código", typeof(string));
                dt.Columns.Add("Modelo", typeof(string));
                dt.Columns.Add("Usuario", typeof(string));
                dt.Columns.Add("Observação", typeof(string));
                dt.Columns.Add("PN", typeof(string));
                dt.Columns.Add("Anexo", typeof(string));

                for (int i = 0; i < oRecSetBuscarAprov.RecordCount; i++)
                {
                    var DadosAProv = new DadosAProv()
                    {
                        Codigo = oRecSetBuscarAprov.Fields.Item("Codigo").Value.ToString(),
                        Modelo = oRecSetBuscarAprov.Fields.Item("Modelo").Value.ToString(),
                        Usuario = oRecSetBuscarAprov.Fields.Item("Usuario").Value.ToString(),
                        Obervacao = oRecSetBuscarAprov.Fields.Item("Observacao").Value.ToString(),
                        PN = oRecSetBuscarAprov.Fields.Item("PN").Value.ToString(),
                        Total =oRecSetBuscarAprov.Fields.Item("Total").Value.ToString(),
                        Status = oRecSetBuscarAprov.Fields.Item("Status").Value.ToString(),
                        Anexo = oRecSetBuscarAprov.Fields.Item("Anexo").Value.ToString(),

                    };

                    //Preencher o grid
                    dt.Rows.Add(DadosAProv.Status, DadosAProv.Total, DadosAProv.Codigo, DadosAProv.Modelo, DadosAProv.Usuario, DadosAProv.Obervacao, DadosAProv.PN, DadosAProv.Anexo);

                    dgvDados.DataSource = dt;
                    oRecSetBuscarAprov.MoveNext();
                }

            }
        }

        private void Pesquisar(string user)
        {
            squery = @"select 
                                convert(nvarchar(100),c1.trgtPath) + '\' + c1.[FILENAME] + '.' + c1.FileExt [Anexo]
                                ,rf.DocNum Codigo
                                , tm.[Name] Modelo
                                ,sr.U_NAME [Usuario]
                                ,dd.Remarks [Observacao]
                                ,rf.CardName PN
                                ,FORMAT( rf.DocTotal, 'C', 'pt-br') Total

                                 ,case 
                                when status = 'W' then'Pendente'
                                when status = 'Y' then 'Aprovado'
                                when status = 'N'then 'Rejeitado'
                                when status = 'P' then  'Gerado'
                                when status = 'A' then ' Gerado pelo Autorizador'
                                when status = 'C' then  'Canceled' end as Status

                                from
                                owdd dd 
								inner join owtm tm on tm.WtmCode = dd.WtmCode
								inner join ousr sr on sr.USERID = dd.UserSign
								left join OWST st  on st.WstCode= dd.CurrStep 
								left join WST1 st1 on st1.WstCode = st.WstCode
								inner join ODRF rf on rf.DocEntry = dd.DraftEntry and rf.ObjType in ('22','204') 
								left join atc1 c1 on c1.AbsEntry = rf.AtcEntry
								inner join OUSR on OUSR.USERID = st1.UserID
                                where tm.Active = 'Y'  and dd.ProcesStat in ('Y','W') and rf.DocStatus <> 'C' and OUSR.USER_CODE = '" + user + "' " + status;



        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialog = new DialogResult();
            dialog = MessageBox.Show("Deseja mesmo encerrar?", "Alerta!", MessageBoxButtons.YesNo);

            if (dialog == DialogResult.Yes)
            {
                ConectaSAP.oCompany.Disconnect();
                Application.Exit();
            }
        }
    }
}
