namespace Skyline.DataMiner.Library.Common.InterAppCalls.CallBulk
{
    using Skyline.DataMiner.Library.Common;
    using Skyline.DataMiner.Library.Common.Attributes;
    using Skyline.DataMiner.Library.Common.InterAppCalls.CallSingle;
    using Skyline.DataMiner.Library.Common.InterAppCalls.Shared;
    using Skyline.DataMiner.Library.Common.Serializing;
    using Skyline.DataMiner.Library.Common.Subscription.Waiters.InterApp;
    using Skyline.DataMiner.Net;

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;

    [DllImport("System.Runtime.Serialization.dll")]
    internal class InterAppCall : IInterAppCall
    {
        public InterAppCall(string guid)
        {
            if (String.IsNullOrWhiteSpace(guid))
            {
                throw new ArgumentNullException("guid", "Identifier should not be empty or null.");
            }

            Guid = guid;
            Messages = new Messages(this);
        }

        public InterAppCall()
        {
            Guid = System.Guid.NewGuid().ToString();
            Messages = new Messages(this);
        }

		public string Guid { get; set; }

        public Messages Messages { get; private set; }

        public DateTime ReceivingTime { get; set; }

        public ReturnAddress ReturnAddress { get; set; }

        public DateTime SendingTime { get; private set; }

        public Source Source { get; set; }

        public void Send(IConnection connection, int agentId, int elementId, int parameterId, List<Type>knownTypes)
        {
			var defaultSerializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);
			Send(connection, agentId, elementId, parameterId, defaultSerializer);
		}

        public void Send(IConnection connection, int agentId, int elementId, int parameterId, ISerializer serializer)
        {
			DmsElementId destination = new DmsElementId(agentId, elementId);
			BubbleDownReturn();
			SendToElement(connection, destination, parameterId, serializer);
		}

        public IEnumerable<Message> Send(IConnection connection, int agentId, int elementId, int parameterId, TimeSpan timeout, List<Type>knownTypes)
        {			
			var defaultSerializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);
			return Send(connection, agentId, elementId, parameterId, timeout, defaultSerializer);
		}

        public IEnumerable<Message> Send(IConnection connection, int agentId, int elementId, int parameterId, TimeSpan timeout, ISerializer serializer)
        {
			if (ReturnAddress != null)
			{
				BubbleDownReturn();

				using (MessageWaiter waiter = new MessageWaiter(new ConnectionCommunication(connection), serializer, serializer, Messages.ToArray()))
				{
					DmsElementId destination = new DmsElementId(agentId, elementId);
					SendToElement(connection, destination, parameterId, serializer);
					foreach (var returnedMessage in waiter.WaitNext(timeout))
					{
						yield return returnedMessage;
					}
				}
			}
			else
			{
				throw new InvalidOperationException("Call is missing ReturnAddress, either add a ReturnAddress or send without defined timeout.");
			}
        }

        private void BubbleDownReturn()
        {
            foreach (var message in Messages)
            {
                if (message.Source == null && Source != null) message.Source = Source;
                if (ReturnAddress != null) message.ReturnAddress = ReturnAddress;
            }
        }

        private void SendToElement(IConnection connection, DmsElementId destination, int parameterId, ISerializer internalSerializer)
        {
            IDms thisDms = connection.GetDms();
            var element = thisDms.GetElement(destination);

            if (element.State == ElementState.Active)
            {
                var parameter = element.GetStandaloneParameter<string>(parameterId);
                SendingTime = DateTime.Now;
                string value = internalSerializer.SerializeToString(this);
                parameter.SetValue(value);
            }
            else
            {
                throw new InvalidOperationException("Could not send message to element " + element.Name + "(" + element.DmsElementId + ")" + " with state " + element.State);
            }
        }
    }
}