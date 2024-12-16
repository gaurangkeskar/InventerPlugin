using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using ShowProperties;

namespace InventerPlugin
{


    [Guid("c322e19b-5e98-46c5-bfe1-fe8cd2e4c173")]
    [ComVisible(true)]
    public class MyInventorAddin : ApplicationAddInServer
    {
        private Inventor.Application _inventorApp;
        private ButtonDefinition _buttonDef;
        private RibbonPanel _ribbonPanel;
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inventorApp = addInSiteObject.Application;
           
            Ribbon ribbon = _inventorApp.UserInterfaceManager.Ribbons["Part"];

            RibbonTab ribbonTab = ribbon.RibbonTabs["id_TabTools"];
            _ribbonPanel = ribbonTab.RibbonPanels.Add("My Custom Panel", "My Custom Panel", "", "");

            _buttonDef = _inventorApp.CommandManager.ControlDefinitions.AddButtonDefinition("Show Properties","Properties", CommandTypesEnum.kNonShapeEditCmdType);

            _ribbonPanel.CommandControls.AddButton(_buttonDef);

            _buttonDef.OnExecute += ButtonDef_OnExecute;
        }

        private void ButtonDef_OnExecute(NameValueMap context)
        {
            if (_inventorApp.ActiveDocument is Document activeDocument)
            {              
                Dictionary<string, string> properties = new Dictionary<string, string>();
                foreach (PropertySet propertySet in activeDocument.FilePropertySets)
                {
                    foreach (Property property in propertySet)
                    {
                        properties[property.Name] = property.Value?.ToString() ?? "N/A";
                    }
                }

                MainWindow window = new MainWindow();
                window.PopulateProperties(properties);
                window.Show();

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No active document found!", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        public void Deactivate()
        {
            // Cleanup resources
            _inventorApp = null;
        }

        public void ExecuteCommand(int commandId)
        {
            // This method can remain empty
        }

        public object Automation => null;
    }
}
