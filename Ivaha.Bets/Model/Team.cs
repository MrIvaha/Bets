using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ivaha.Bets.Model
{
    public  class   Team
    {
        public  string                          Name                { get; set; }

        public              List<Team>          Winners             { get; set; }               =   new List<Team>();
        public              List<Team>          Losers              { get; set; }               =   new List<Team>();
        public              List<Team>          Tied                { get; set; }               =   new List<Team>();

        public              string[]            WinnersAndLosers    { get; protected set; }     =   new string[0];
        public              string[]            OnlyWinners         { get; protected set; }     =   new string[0];
        public              string[]            OnlyLosers          { get; protected set; }     =   new string[0];
        public              string[]            OnlyTied            { get; protected set; }     =   new string[0];

        public              bool                IsBettable                                      =>  OnlyWinners.Length > 0 || OnlyLosers.Length > 0 || OnlyTied.Length > 0;

        protected   virtual List<Func<string[]>>MakeListFuncs       { get; }

        public                                  Team                ()
        {
            MakeListFuncs   =   new List<Func<string[]>>{ MakeWinnersAndLosers, MakeOnlyWinners, MakeOnlyLosers, MakeOnlyTied };
        }
        public                                  Team                (string name) : this ()     =>  Name    =   name;
        public     override string              ToString            ()  =>  Name;

        public      virtual string[]            MakeWinnersAndLosers()  =>  WinnersAndLosers    =   Winners?.Where(t => Losers?.Contains(t) ?? false).Select(t => t.Name).ToArray() ?? new string[0];
        public      virtual string[]            MakeOnlyWinners     ()  =>  OnlyWinners         =   Winners?.Where(t => !(Losers?.Contains(t) ?? false)).Select(t => t.Name).ToArray()  ?? new string[0];
        public      virtual string[]            MakeOnlyLosers      ()  =>  OnlyLosers          =   Losers?.Where(t => !(Winners?.Contains(t) ?? false)).Select(t => t.Name).ToArray()  ?? new string[0];
        public      virtual string[]            MakeOnlyTied        ()  =>  OnlyTied            =   Tied?/*.Where(t => !Losers.Contains(t) && !Winners.Contains(t))*/.Select(t => t.Name).ToArray()  ?? new string[0];
        public              void                MakeAllLists        ()  =>  MakeListFuncs.ForEach(f => f.Invoke());
    }
}
