using Castle.Core.Logging;
using JBC.COLETOR.Model;
using JBC.COLETOR.services;
using JBC.Framework.Form;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JBC.COLETOR.FormClass.systemform
{
    [FormAttribute("85", "JBC.COLETOR.resources.frmPicking.srf")]
    public class frmPicking : JBCSystemFormBase
    {
        public SAPbobsCOM.Company oCompany { get; set; }
        public SAPbouiCOM.Application oApplication { get; set; }
        public JBC.Framework.DAO.BusinessOneDAO B1DAO { get; set; }
        public ILogger Log { get; set; }
        private SAPbouiCOM.IForm _oForm;

        private SAPbouiCOM.Button _ButtonColetar;
        private SAPbouiCOM.Button _ButtonCancelar;

        private string errMsg = "";
        private int errCode = 0;

        JBCCOLETORService _JBCCOLETORService;
        private JBCCOLETORService oJBCCOLETORService
        {
            get
            {
                if (_JBCCOLETORService == null)
                {
                    _JBCCOLETORService = new JBCCOLETORService(this.B1DAO, this.Log, this.oCompany);
                }
                return _JBCCOLETORService;
            }
        }

        public override void OnInitializeComponent()
        {            
            if (this._oForm != null)
            {
                return;
            }
            this._oForm = ((JBC.Framework.Form.JBCFormBase)(this)).UIAPIRawForm;
            this._oForm.Title = "Lista de Picking - Claudino Soluções em Software";
            this._ButtonCancelar = ((SAPbouiCOM.Button)(this.GetItem("4").Specific));
            this._ButtonColetar = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            SetEventHandlers();
            PosicaoBotaoColetar();
        }

        private void PosicaoBotaoColetar()
        {
            this._ButtonColetar.Item.Top = this._ButtonCancelar.Item.Top;
            this._ButtonColetar.Item.Left = this._ButtonCancelar.Item.Left + this._ButtonCancelar.Item.Width + 10;
            this._ButtonColetar.Item.LinkTo = "4";
        }
        private void SetEventHandlers(bool exit = false)
        {
            if (!exit)
            {
                this._ButtonColetar.PressedAfter += _ButtonColetar_PressedAfter;
            }
            else
            {
                this._ButtonColetar.PressedAfter -= _ButtonColetar_PressedAfter;
            }
            
        }

        protected internal virtual void _ButtonColetar_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            List<PickingItem> lstPickingItems = oJBCCOLETORService.GetPickingItems(_oForm.DataSources.DBDataSources.Item("OPKL").GetValue("AbsEntry", 0).ToString());
            int iDocEntry = Convert.ToInt32(_oForm.DataSources.DBDataSources.Item("OPKL").GetValue("AbsEntry", 0).ToString());
            if (lstPickingItems.Count==0)
            {
                Log.WarnFormat("Coleta não Encontrada!");
            }
            else if (_oForm.Mode != BoFormMode.fm_OK_MODE)
            {
                Log.Error("Salve para Continuar...");
            }
            else
            {
                PickLists p = (PickLists)oCompany.GetBusinessObject(BoObjectTypes.oPickLists);
                if (p.GetByKey(iDocEntry))
                {
                    //foreach (PickingItem item in lstPickingItems)
                    //{
                        for (int i = 0; i < p.Lines.Count; i++)
                        {
                            p.Lines.SetCurrentLine(i);
                            p.Lines.PickedQuantity = p.Lines.ReleasedQuantity;
                        }
                    //}
                    if (p.Update()==0)
                    {
                        oApplication.ActivateMenuItem("1304");
                        Log.InfoFormat("Coleta Efetuada com Sucesso!!!");
                    }
                    else
                    {
                        oCompany.GetLastError(out errCode, out errMsg);
                        errMsg = string.Format("[{0}] {1}", errCode, errMsg);
                        Log.Error(errMsg);
                    }
                }
                else
                {

                }

            }
        }

        protected override void OnFormCloseAfter(SBOItemEventArg pVal)
        {
            SetEventHandlers(true);
        }
        protected override void OnFormResizeBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            try
            {
                this._oForm.Freeze(true);
                PosicaoBotaoColetar();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }
            finally
            {
                this._oForm.Freeze(false);
            }
            BubbleEvent = true;
        }
        protected override void OnFormResizeAfter(SBOItemEventArg pVal)
        {
            try
            {
                this._oForm.Freeze(true);
                PosicaoBotaoColetar();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }
            finally
            {
                if (this._oForm != null)
                {
                    this._oForm.Freeze(false);
                }

            }
        }
    }
}
