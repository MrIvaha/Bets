using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace Ivaha.Bets.ViewModel
{
    class   MagicAttribute  : Attribute { }
    class   NoMagicAttribute: Attribute { }

    [Magic]
    public abstract class   ViewModelBase : INotifyPropertyChanged, IDisposable
    {
        public      event   PropertyChangedEventHandler 
                                    PropertyChanged;

        public      Window          MainControl             { get; set; }

        Type                        CommandOwnerType;
        Dictionary<Type, List<CommandBinding>>
                                    regCommands             =   new Dictionary<Type, List<CommandBinding>>();

        protected                   ViewModelBase           (){ }
        protected                   ViewModelBase           (Window mainControl, Type commandOwnerType)
        {
            MainControl         =   mainControl;
            CommandOwnerType    =   commandOwnerType;
        }
        protected   virtual void    RaisePropertyChanged    (string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
        public              void    RegCommand              (ICommand command, ExecutedRoutedEventHandler executed = null, CanExecuteRoutedEventHandler canExecute = null)
        {
            var cmdBinding  =   new CommandBinding(command, executed, canExecute);
            MainControl?.CommandBindings.Add(cmdBinding);

            if (command is RoutedCommand routedCmd)
            {
                if (!regCommands.ContainsKey(routedCmd.OwnerType))
                    regCommands.Add(routedCmd.OwnerType, new List<CommandBinding>());

                regCommands[routedCmd.OwnerType].Add(cmdBinding);
            }
        }
        public              void    UnregCommand            (Type ownerType)
        {
            if (MainControl == null || !regCommands.ContainsKey(ownerType))
                return;

            foreach (var cmd in regCommands[ownerType])
                MainControl.CommandBindings.Remove(cmd);
        }
        public              void    Dispose                 ()
        {
            UnregCommand(CommandOwnerType);
        }
    }
}
