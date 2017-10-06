using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelToCsvWpfBased
{
    internal sealed class MainWindowCommand : ICommand
    {
        private Action myExecuteAction;
        private Func<bool> myCanExcecutePredicate;

        public MainWindowCommand(Action executeAction)
            : this(executeAction, null)
        {
        }

        public MainWindowCommand(Action executeAction, Func<bool> canExcecutePredicate)
        {
            if (null == executeAction)
            {
                throw new ArgumentNullException("executeAction");
            }

            myExecuteAction = executeAction;
            myCanExcecutePredicate = canExcecutePredicate;
        }

        public bool CanExecute(object parameter)
        {
            return myCanExcecutePredicate == null ? true : myCanExcecutePredicate.Invoke();
        }

        public void Execute(object parameter)
        {
            InvokeExcecuting();

            InvokeExcecute();

            InvokeExcecuted();
        }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                if (myCanExcecutePredicate != null)
                {
                    CommandManager.RequerySuggested += value;
                }
            }
            remove
            {
                if (myCanExcecutePredicate != null)
                {
                    CommandManager.RequerySuggested -= value;
                }
            }
        }

        /// <summary>
        /// Occurs when the command is about to execute.
        /// </summary>
        public event EventHandler<MessageNotificationEventArgs> Excecuting;

        /// <summary>
        /// Occurs when the command executed.
        /// </summary>
        public event EventHandler<MessageNotificationEventArgs> Excecuted;

        private void InvokeExcecute()
        {
            var handler = myExecuteAction;
            if (handler != null)
            {
                handler();
            }
        }

        private void InvokeExcecuting()
        {
            var handler = Excecuting;
            if (handler != null)
            {
                handler(this, new MessageNotificationEventArgs { Message = "Are you sure, you want to convert?", MessageType = MessageType.Confirmation });
            }
        }

        private void InvokeExcecuted()
        {
            var handler = Excecuted;
            if (handler != null)
            {
                handler(this, new MessageNotificationEventArgs { Message = "File(s) converted successfully", MessageType = MessageType.Notification });
            }
        }
    }

    internal sealed class MessageNotificationEventArgs : EventArgs
    {
        public string Message { get; set; }
        public MessageType MessageType { get; set; }
    }
}
