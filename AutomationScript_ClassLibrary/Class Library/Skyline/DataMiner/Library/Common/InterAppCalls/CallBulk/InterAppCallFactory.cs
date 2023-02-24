namespace Skyline.DataMiner.Library.Common.InterAppCalls.CallBulk
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Skyline.DataMiner.Library.Common;
	using Skyline.DataMiner.Library.Common.InterAppCalls.CallSingle;
	using Skyline.DataMiner.Library.Common.Serializing;
	using Skyline.DataMiner.Net;

	/// <summary>
	/// Factory class that can create inter-app calls.
	/// </summary>
	public static class InterAppCallFactory
	{
		/// <summary>
		/// Creates an inter-app call from the specified string.
		/// </summary>
		/// <param name="rawData">The serialized raw data.</param>
		/// <returns>An inter-app call.</returns>
		/// <param name="serializer">Serializer to use.</param>
		/// <exception cref="ArgumentNullException"><paramref name="rawData"/> is empty or null.</exception>
		/// <exception cref="ArgumentException">Format of <paramref name="rawData"/> is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRaw(string rawData, ISerializer serializer)
		{
			if (String.IsNullOrWhiteSpace(rawData)) throw new ArgumentNullException("rawData");
			var returnedResult = serializer.DeserializeFromString<InterAppCall>(rawData);
			returnedResult.ReceivingTime = DateTime.Now;
			return returnedResult;
		}

		/// <summary>
		/// Creates an inter-app call from the specified string.
		/// </summary>
		/// <param name="rawData">The serialized raw data.</param>
		/// <returns>An inter-app call.</returns>
		/// <param name="knownTypes">A list of known message types.</param>
		/// <exception cref="ArgumentNullException"><paramref name="rawData"/> is empty or null.</exception>
		/// <exception cref="ArgumentException">Format of <paramref name="rawData"/> is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRaw(string rawData, IEnumerable<Type> knownTypes)
		{
			if (String.IsNullOrWhiteSpace(rawData)) throw new ArgumentNullException("rawData");
			var serializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);

			return CreateFromRaw(rawData, serializer);
		}

		/// <summary>
		/// Creates an inter-app call from the specified string.
		/// </summary>
		/// <param name="rawData">The serialized raw data.</param>
		/// <returns>An inter-app call.</returns>
		/// <param name="interAppSerializer">Serializer to use for InterAppCall.</param>
		/// <param name="messageSerializer">Serializer to use for Message.</param>
		/// <exception cref="ArgumentNullException"><paramref name="rawData"/> is empty or null.</exception>
		/// <exception cref="ArgumentException">Format of <paramref name="rawData"/> is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRawAndAcceptMessage(string rawData, ISerializer interAppSerializer, ISerializer messageSerializer)
		{
			if (String.IsNullOrWhiteSpace(rawData))
			{
				throw new ArgumentNullException("rawData");
			}

			IInterAppCall returnedResult;
			try
			{
				returnedResult = interAppSerializer.DeserializeFromString<InterAppCall>(rawData);
			}
			catch (Exception)
			{
				Message message;
				if (MessageFactory.TryCreateFromRaw(rawData, out message, messageSerializer))
				{
					returnedResult = CreateNew();
					returnedResult.Messages.Add(message);
					returnedResult.Source = message.Source;
					returnedResult.ReturnAddress = message.ReturnAddress;
				}
				else
				{
					throw;
				}
			}

			returnedResult.ReceivingTime = DateTime.Now;
			return returnedResult;
		}

		/// <summary>
		/// Creates an inter-app call from the specified string.
		/// </summary>
		/// <param name="rawData">The serialized raw data.</param>
		/// <returns>An inter-app call.</returns>
		/// <param name="knownTypes">A list of known message types.</param>
		/// <exception cref="ArgumentNullException"><paramref name="rawData"/> is empty or null.</exception>
		/// <exception cref="ArgumentException">Format of <paramref name="rawData"/> is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRawAndAcceptMessage(string rawData, IEnumerable<Type> knownTypes)
		{
			if (String.IsNullOrWhiteSpace(rawData))
			{
				throw new ArgumentNullException("rawData");
			}

			ISerializer interAppSerializer;
			ISerializer messageSerializer;

			IInterAppCall returnedResult;
			try
			{
				interAppSerializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);
				returnedResult = interAppSerializer.DeserializeFromString<InterAppCall>(rawData);
				if (!returnedResult.Messages.Any()) throw new InvalidCastException();
			}
			catch (Exception)
			{
				messageSerializer = SerializerFactory.CreateInterAppSerializer(typeof(Message), knownTypes);

				Message message;
				if (MessageFactory.TryCreateFromRaw(rawData, out message, messageSerializer))
				{
					returnedResult = CreateNew();
					returnedResult.Messages.Add(message);
					returnedResult.Source = message.Source;
					returnedResult.ReturnAddress = message.ReturnAddress;
				}
				else
				{
					throw;
				}
			}

			returnedResult.ReceivingTime = DateTime.Now;
			return returnedResult;
		}

		/// <summary>
		/// Creates an inter-app call from the contents of the specified parameter.
		/// </summary>
		/// <param name="connection">The raw SLNet connection.</param>
		/// <param name="agentId">The source DataMiner Agent ID.</param>
		/// <param name="elementId">The source element ID.</param>
		/// <param name="parameterId">The source parameter ID.</param>
		/// <param name="serializer">Serializer to use.</param>
		/// <returns>An inter-app call.</returns>
		/// <exception cref="ArgumentException">The format of the content of the specified parameter is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRemote(IConnection connection, int agentId, int elementId, int parameterId, ISerializer serializer)
		{
			IDms thisDms = connection.GetDms();
			var element = thisDms.GetElement(new DmsElementId(agentId, elementId));
			var parameter = element.GetStandaloneParameter<string>(parameterId);
			var returnedResultRaw = parameter.GetValue();

			return CreateFromRaw(returnedResultRaw, serializer);
		}

		/// <summary>
		/// Creates an inter-app call from the contents of the specified parameter.
		/// </summary>
		/// <param name="connection">The raw SLNet connection.</param>
		/// <param name="agentId">The source DataMiner Agent ID.</param>
		/// <param name="elementId">The source element ID.</param>
		/// <param name="parameterId">The source parameter ID.</param>
		/// <param name="knownTypes">A list of known message types.</param>
		/// <returns>An inter-app call.</returns>
		/// <exception cref="ArgumentException">The format of the content of the specified parameter is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRemote(IConnection connection, int agentId, int elementId, int parameterId, IEnumerable<Type> knownTypes)
		{
			IDms thisDms = connection.GetDms();
			var element = thisDms.GetElement(new DmsElementId(agentId, elementId));
			var parameter = element.GetStandaloneParameter<string>(parameterId);
			var returnedResultRaw = parameter.GetValue();
			var serializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);

			return CreateFromRaw(returnedResultRaw, serializer);
		}

		/// <summary>
		/// Creates an inter-app call from the contents of the specified parameter.
		/// </summary>
		/// <param name="connection">The raw SLNet connection.</param>
		/// <param name="agentId">The source DataMiner Agent ID.</param>
		/// <param name="elementId">The source element ID.</param>
		/// <param name="parameterId">The source parameter ID.</param>
		/// <param name="interAppSerializer">Serializer to use for InterAppCall.</param>
		/// <param name="messageSerializer">Serializer to use for Message.</param>
		/// <returns>An inter-app call.</returns>
		/// <exception cref="ArgumentException">The format of the content of the specified parameter is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRemoteAndAcceptMessage(IConnection connection, int agentId, int elementId, int parameterId, ISerializer interAppSerializer, ISerializer messageSerializer)
		{
			IDms thisDms = connection.GetDms();
			IDmsElement element = thisDms.GetElement(new DmsElementId(agentId, elementId));
			IDmsStandaloneParameter<string> parameter = element.GetStandaloneParameter<string>(parameterId);
			string returnedResultRaw = parameter.GetValue();

			return CreateFromRawAndAcceptMessage(returnedResultRaw, interAppSerializer, messageSerializer);
		}

		/// <summary>
		/// Creates an inter-app call from the contents of the specified parameter.
		/// </summary>
		/// <param name="connection">The raw SLNet connection.</param>
		/// <param name="agentId">The source DataMiner Agent ID.</param>
		/// <param name="elementId">The source element ID.</param>
		/// <param name="parameterId">The source parameter ID.</param>
		/// <param name="knownTypes">A list of known message types.</param>
		/// <returns>An inter-app call.</returns>
		/// <exception cref="ArgumentException">The format of the content of the specified parameter is invalid and deserialization failed.</exception>
		public static IInterAppCall CreateFromRemoteAndAcceptMessage(IConnection connection, int agentId, int elementId, int parameterId, IEnumerable<Type> knownTypes)
		{
			IDms thisDms = connection.GetDms();
			IDmsElement element = thisDms.GetElement(new DmsElementId(agentId, elementId));
			IDmsStandaloneParameter<string> parameter = element.GetStandaloneParameter<string>(parameterId);
			string returnedResultRaw = parameter.GetValue();
			var interAppSerializer = SerializerFactory.CreateInterAppSerializer(typeof(InterAppCall), knownTypes);
			var messageSerializer = SerializerFactory.CreateInterAppSerializer(typeof(Message), knownTypes);
			return CreateFromRawAndAcceptMessage(returnedResultRaw, interAppSerializer, messageSerializer);
		}

		/// <summary>
		/// Creates a blank inter-app call.
		/// </summary>
		/// <returns>An inter-app call.</returns>
		public static IInterAppCall CreateNew()
		{
			return new InterAppCall();
		}
	}
}