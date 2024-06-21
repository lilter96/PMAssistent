using Microsoft.Office.Interop.MSProject;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ProjectTaskProcessor
{
    [ComVisible(true)]
    public class AssigneeRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private CancellationTokenSource cancellationTokenSource;

        private const string AssigneeLastName = "Гацуков";
        private const string TaskContains = "Исполнитель";
        private const int TaskDelay = 150;

        public AssigneeRibbon()
        {
        }

        public async void OnAssigneeSelected(Office.IRibbonControl control)
        {
            var application = Globals.ThisAddIn.Application;
            var activeProject = application.ActiveProject;

            ProgressForm progressForm = new ProgressForm(activeProject.Tasks.Count);

            cancellationTokenSource = new CancellationTokenSource();

            try
            {
                progressForm.Show();

                int affectedTasks = await ProcessTasksAsync(activeProject, progressForm, cancellationTokenSource.Token);

                MessageBox.Show($"Исполнитель {AssigneeLastName} был добавлен к {affectedTasks} задачам.", "Завершено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Операция была отменена.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressForm.Close();
                cancellationTokenSource.Dispose();
            }
        }

        private async Task<int> ProcessTasksAsync(Project activeProject, ProgressForm progressForm, CancellationToken token)
        {
            return await System.Threading.Tasks.Task.Run(async () =>
            {
                int affectedTasks = 0;

                var resource = activeProject.Resources.Cast<Resource>().FirstOrDefault(r => r.Name.Contains(AssigneeLastName));
                if (resource == null)
                {
                    resource = activeProject.Resources.Add(AssigneeLastName);
                }

                foreach (Microsoft.Office.Interop.MSProject.Task task in activeProject.Tasks)
                {
                    token.ThrowIfCancellationRequested();

                    // Asynchronous delay for work simulation
                    await System.Threading.Tasks.Task.Delay(TaskDelay, token);

                    if (task.Name.Contains(TaskContains))
                    {
                        ClearAssignments(task);
                        task.Assignments.Add(task.ID, resource.ID);
                        affectedTasks++;
                    }

                    progressForm.Invoke(new Action(() => progressForm.UpdateProgress()));
                }

                return affectedTasks;
            }, token);
        }

        private void ClearAssignments(Microsoft.Office.Interop.MSProject.Task task)
        {
            foreach (Assignment assignment in task.Assignments)
            {
                assignment.Delete();
            }
        }

        public void CancelOperation()
        {
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Cancel();
            }
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ProjectTaskProcessor.AssigneeRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
