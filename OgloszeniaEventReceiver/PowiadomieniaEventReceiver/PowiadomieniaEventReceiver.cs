using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace PublishAnnouncementEventReceiver.PowiadomieniaEventReceiver
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class PowiadomieniaEventReceiver : SPItemEventReceiver
    {

        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            RunAssociatedWorkflow_OnChange(properties);
        }

       public override void ItemUpdated(SPItemEventProperties properties)
       {
           base.ItemUpdated(properties);

           RunAssociatedWorkflow_OnChange(properties);
       }

       private void RunAssociatedWorkflow_OnChange(SPItemEventProperties properties)
       {
           this.EventFiringEnabled = false;

           
           /// opcja 1
           
           //// put here the ID of your workflow as it is in the workflow feature definition
           //Guid workflowDefinitionID = new Guid("A3A980FD-A6DE-45A7-B68F-A54EFA85A451");

           //SPListItem item = properties.ListItem;
           ////if ((int)item[SPBuiltInFieldId._ModerationStatus] == (int)SPModerationStatusType.Approved)
           ////{
           //    SPWorkflowManager objWorkflowManager = item.ParentList.ParentWeb.Site.WorkflowManager;
           //    SPWorkflowAssociation wAssoc = item.ParentList.WorkflowAssociations.Cast<SPWorkflowAssociation>().FirstOrDefault(wa => wa.BaseId == workflowDefinitionID);
           //    if (wAssoc != null)
           //    {
           //        objWorkflowManager.StartWorkflow(item, wAssoc, wAssoc.AssociationData, true);
           //    }
           ////}

           /// opcja 2

           //SPListItem item = properties.ListItem;

           //SPWorkflowManager wfManager =
           //    item.ParentList.ParentWeb.Site.WorkflowManager;
           //SPWorkflowAssociationCollection wfAssociationCollection =
           //    item.ParentList.WorkflowAssociations;

           //foreach (SPWorkflowAssociation wfAssociation in wfAssociationCollection)
           //{
           //    //identyfikator BaseId workflow weź z konfiguracji WF w SPD lub z pliku AssemblyInfo.cs w VS
           //    if (wfAssociation.BaseId == new Guid("A3A980FD-A6DE-45A7-B68F-A54EFA85A451"))
           //    {
           //        try
           //        {
           //             wfManager.StartWorkflow(item, wfAssociation, wfAssociation.AssociationData, true);
           //        }
           //        catch (Exception exp)
           //        {
           //            throw;
           //        }
                   
           //        break;
           //    }
           //}


           /// opcja 3
           
           SPListItem item = properties.ListItem;
           StartWorkflow(item, "Powiadomienie.OnCreate");


           this.EventFiringEnabled = true;
       }

        #region Helpers

       private static void StartWorkflow(SPListItem listItem, string workflowName)
       {
           SPWorkflowManager manager = listItem.Web.Site.WorkflowManager;
           SPWorkflowAssociationCollection objWorkflowAssociationCollection = listItem.ParentList.WorkflowAssociations;
           foreach (SPWorkflowAssociation objWorkflowAssociation in objWorkflowAssociationCollection)
           {
               if (String.Compare(objWorkflowAssociation.Name, workflowName, true) == 0)
               {

                   //We found our workflow association that we want to trigger.

                   //Replace the workflow_GUID with the GUID of the workflow feature that you
                   //have deployed.

                   try
                   {
                       manager.StartWorkflow(listItem, objWorkflowAssociation, objWorkflowAssociation.AssociationData, true);
                       //The above line will start the workflow...
                   }
                   catch (Exception)
                   {}
                   
                   
                   break;
               }
           }
       }

        #endregion

    }
}
