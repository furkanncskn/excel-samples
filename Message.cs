using System.Windows.Forms;

namespace Excel
{
    public static class Message
    {
        public static void ShowMessage(string message, MessageType type)
        {
            switch (type)
            {
                case MessageType.Error:
                    MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case MessageType.Information:
                    MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case MessageType.Warning:
                    MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case MessageType.Abort:
                    MessageBox.Show(message, "Abort", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                default:
                    break;
            }
        }

        public enum MessageType
        {
            Error,
            Information,
            Warning,
            Abort
        }
    }
}
