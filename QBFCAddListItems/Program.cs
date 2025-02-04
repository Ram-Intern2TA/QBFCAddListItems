using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Interop.QBFC16;

namespace QuickBooksItemList
{
    class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.DoItemNonInventoryAdd();
        }

        public void DoItemNonInventoryAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemNonInventoryAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                WalkItemNonInventoryAddRs(responseMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }

            void BuildItemNonInventoryAddRq(IMsgSetRequest requestMsgSet)
            {
                try
                {
                    IItemServiceAdd ItemServiceAddRq = requestMsgSet.AppendItemServiceAddRq();
                    ItemServiceAddRq.Name.SetValue("Sample1#");
                    ItemServiceAddRq.IsActive.SetValue(true);
                   
                    ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(15.65);

                    ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue("Revenue");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error in request build: {ex.Message}");
                }
            }

            void WalkItemNonInventoryAddRs(IMsgSetResponse responseMsgSet)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0)
                {
                    Console.WriteLine("Item added successfully.");
                }
                else
                {
                    Console.WriteLine($"Error adding item: {response.StatusMessage}");
                }
            }
        }
    }
}
