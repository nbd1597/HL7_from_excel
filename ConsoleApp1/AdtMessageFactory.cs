using System;
using NHapi.Base.Model;

namespace ExceltoHl7
{
    public class AdtMessageFactory
    {
        //This class incase you want to make different kind of adt message
        public static IMessage CreateMessage(string messageType)
        {
            //This patterns enables you to build other message types 
            if (messageType.Equals("A01"))
            {
                return new ExcelParser().ReadTemplte();
            }

            //if other types of ADT messages are needed, then implement your builders here
            throw new ArgumentException($"'{messageType}' is not supported yet.");
        }
    }
}