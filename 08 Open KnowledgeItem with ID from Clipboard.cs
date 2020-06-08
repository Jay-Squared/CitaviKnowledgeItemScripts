using System;
using System.IO;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

public static class CitaviMacro
{
	public static void Main()
	{
		Project project = Program.ActiveProjectShell.Project;		
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;		
		
		List<KnowledgeItem> allKnowledgeItems = project.AllKnowledgeItems.ToList();
		List<KnowledgeItem> foundKnowledgeItems = new List<KnowledgeItem>();
		
		using (var reader = new StringReader(Clipboard.GetText()))
		{
		    for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
		    {
		        KnowledgeItem knowledgeItem = allKnowledgeItems.Where(k => k.Id.ToString() == line).FirstOrDefault();
				if (knowledgeItem == null) continue;
				foundKnowledgeItems.Add(knowledgeItem);
				MessageBox.Show(knowledgeItem.CoreStatement);
		    }
		}
		
		if (mainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
		{
			foreach (KnowledgeItem k in foundKnowledgeItems)
			{				
				Program.ActiveProjectShell.ShowKnowledgeItemFormForExistingItem(mainForm, k);
			}
		}
		else if (mainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
		{
			if(foundKnowledgeItems.Count > 0)
			{
				KnowledgeItemFilter filter = new KnowledgeItemFilter(foundKnowledgeItems, "Knowledge Items Selected in Word", false);
				Program.ActiveProjectShell.PrimaryMainForm.KnowledgeOrganizerFilterSet.Filters.ReplaceBy(new List<KnowledgeItemFilter> { filter });
			}
		}		
	}
}