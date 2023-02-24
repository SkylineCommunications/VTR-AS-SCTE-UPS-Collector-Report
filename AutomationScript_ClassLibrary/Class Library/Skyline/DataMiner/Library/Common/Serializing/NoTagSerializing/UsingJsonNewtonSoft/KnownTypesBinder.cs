namespace Skyline.DataMiner.Library.Common.Serializing.NoTagSerializing.UsingJsonNewtonSoft
{
	using Newtonsoft.Json.Serialization;

	using Skyline.DataMiner.Library.Common.Attributes;

	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Reflection;

	[DllImport("Newtonsoft.Json.dll")]
	internal class KnownTypesBinder : ISerializationBinder
	{
		private Lazy<string[]> nonUniqueTypeNames;

		public KnownTypesBinder()
		{
		}

		public KnownTypesBinder(IList<Type> knownTypes)
		{
			AddKnownTypes(knownTypes);
		}

		public IList<Type> KnownTypes { get; private set; }

		public void AddKnownTypes(IList<Type> knownTypes)
		{
			if (knownTypes != null)
			{
				KnownTypes = knownTypes;
				nonUniqueTypeNames = new Lazy<string[]>(() => { return KnownTypes.GroupBy(x => x.Name).Where(g => g.Count() > 1).Select(y => y.Key).ToArray(); });
			}
		}

		public void BindToName(Type serializedType, out string assemblyName, out string typeName)
		{
			assemblyName = String.Empty;

			if (serializedType == null)
			{
				throw new ArgumentNullException("serializedType");
			}

			if (KnownTypes != null && KnownTypes.Contains(serializedType) && !nonUniqueTypeNames.Value.Contains(serializedType.Name))
			{
				typeName = serializedType.Name;
			}
			else
			{
				typeName = serializedType.FullName;
			}
		}

#pragma warning disable S3776 // Cognitive Complexity of methods should not be too high

		public Type BindToType(string assemblyName, string typeName)
#pragma warning restore S3776 // Cognitive Complexity of methods should not be too high
		{
			if (typeName == null)
			{
				throw new ArgumentNullException("typeName");
			}

			Type foundType = null;
			bool array = false;
			bool doubleArray = false;

			// To Deal with Double Arrays:
			if (typeName.EndsWith("[][]", StringComparison.Ordinal))
			{
#pragma warning disable S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
#pragma warning disable S3257 // Declarations and initializations should be as concise as possible
				typeName = typeName.TrimEnd(new char[] { '[', ']' });
#pragma warning restore S3257 // Declarations and initializations should be as concise as possible
#pragma warning restore S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
				doubleArray = true;
			}

			// To Deal with Arrays:
			if (typeName.EndsWith("[]", StringComparison.Ordinal))
			{
#pragma warning disable S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
#pragma warning disable S3257 // Declarations and initializations should be as concise as possible
				typeName = typeName.TrimEnd(new char[] { '[', ']' });
#pragma warning restore S3257 // Declarations and initializations should be as concise as possible
#pragma warning restore S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
				array = true;
			}

			if (typeName.StartsWith("System", StringComparison.Ordinal) && foundType == null)
			{
				// MSCORLIB
				var mscorlibAssembly = typeof(Object).Assembly;
				try
				{
					foundType = mscorlibAssembly.GetType(typeName);
				}
				catch
				{
					// Ignore exception in order to see if we find the type. Need to use this for logic unfortunately.
				}

				if (foundType == null)
				{
					// SYSTEM
					var systemAssembly = typeof(System.Uri).Assembly;
					try
					{
						foundType = systemAssembly.GetType(typeName);
					}
					catch
					{
						// Ignore exception in order to see if we find the type. Need to use this for logic unfortunately.
					}
				}

				if (foundType == null)
				{
					// SYSTEM.CORE
					var sysCoreAssembly = typeof(HashSet<>).Assembly;
					try
					{
						foundType = sysCoreAssembly.GetType(typeName);
					}
					catch
					{
						// Ignore exception in order to see if we find the type. Need to use this for logic unfortunately.
					}
				}
			}

			if (KnownTypes != null && foundType == null)
			{
				try
				{
					foundType = KnownTypes.SingleOrDefault(t => t.Name == typeName);
				}
				catch (InvalidOperationException ex)
				{
					throw (new IncorrectDataException("Type Name: " + typeName + " was unique on serialization side but not on deserialization side. Please verify the same KnownTypes List is used on both ends of the communication.", ex));
				}
			}

			if (KnownTypes != null && foundType == null)
			{
				foundType = KnownTypes.SingleOrDefault(t => t.FullName == typeName);
			}

			if (foundType == null)
			{
				// Checks the current assembly.
				foreach (Type t in typeof(KnownTypesBinder).Assembly.GetTypes())
				{
					if (typeName == t.FullName)
					{
						foundType = t;
						break;
					}
				}
			}

			if(foundType == null)
			{
				// Check all known assemblies from knowntypes.

				foreach (var knownType in KnownTypes)
				{
					foundType = knownType.Assembly.GetType(typeName);
					if (foundType != null) break;
				}
			}


			if (foundType == null)
			{
				DefaultSerializationBinder def = new DefaultSerializationBinder();

				if (foundType == null && String.IsNullOrWhiteSpace(assemblyName))
				{
#pragma warning disable S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
					assemblyName = typeof(KnownTypesBinder).Assembly.GetName().Name;
#pragma warning restore S1226 // Method parameters, caught exceptions and foreach variables' initial values should not be ignored
				}

				try
				{
					foundType = def.BindToType(assemblyName, typeName);
				}
				catch
				{
					// Ignore exception in order to see if we find the type. Need to use this for logic unfortunately.
				}

			}

			if (doubleArray)
			{
				foundType = foundType.MakeArrayType().MakeArrayType();
			}
			if (array)
			{
				foundType = foundType.MakeArrayType();
			}

			if(foundType == null)
			{
				throw new InvalidOperationException("Deserialization Failed for type: " + typeName);
			}

			return foundType;
		}
	}
}