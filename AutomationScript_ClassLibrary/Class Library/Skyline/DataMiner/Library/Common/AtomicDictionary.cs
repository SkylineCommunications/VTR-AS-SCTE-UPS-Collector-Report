namespace Skyline.DataMiner.Library.Common
{
	using System;
	using System.Collections.Concurrent;
	using System.Collections.Generic;

	internal class ValueWithLock<T>
	{
		public ValueWithLock(T value, object lockObject) 
		{
			Value = value;
			Lock = lockObject;
		}

		public T Value { get; private set; }

		public object Lock { get; private set; }
	}

	/// <summary>
	/// Performs all actions on values alongside dictionary manipulation in an atomic and thread-safe way.
	/// </summary>
	/// <typeparam name="T">The type of the values in the dictionary.</typeparam>
	internal class AtomicDictionary<T>
	{
		private readonly ConcurrentDictionary<string, ValueWithLock<T>> allValueWithLock = new ConcurrentDictionary<string, ValueWithLock<T>>();
		//private readonly ConcurrentDictionary<string, ValueWithLock<T>> allValues = new ConcurrentDictionary<string, ValueWithLock<T>>();
		//private readonly ConcurrentDictionary<string, object> valueLocks = new ConcurrentDictionary<string, object>();

		/// <summary>
		/// Will perform a thread-safe action using the specific key/value and then remove this key/value. If the key does not exist it will do nothing.
		/// </summary>
		/// <param name="key">The key to remove.</param>
		/// <param name="a">The action to perform before removal.</param>
		/// <returns>A <see cref="Boolean"/> indicating if the removal was a success.</returns>
		public bool ActionAndRemove(string key, Action<T> a)
		{
			if (a == null) return false;

			ValueWithLock<T> myValueAndLock;
			
			if (!allValueWithLock.TryGetValue(key, out myValueAndLock))
			{
				return false;
			}

			object myValueLock = myValueAndLock.Lock;
			T myValue = myValueAndLock.Value;

			lock (myValueLock)
			{
				System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":TryActionAndRemove entered lock " + key);

				a(myValue);

				ValueWithLock<T> removedValue;
				allValueWithLock.TryRemove(key, out removedValue);
			}

			System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":TryActionAndRemove EXIT lock " + key);
			return true;
		}

		/// <summary>
		/// Add the value. If it already exists it will do nothing.
		/// </summary>
		/// <param name="key"></param>
		/// <param name="valueToAdd"></param>
		public void Add(string key, T valueToAdd)
		{
			allValueWithLock.TryAdd(key, new ValueWithLock<T>(valueToAdd, new object()));
		}

		/// <summary>
		/// Tries to add the key/value, or create a new one if it doesn't exist and then perform the thread-safe action indicated with that key/value.
		/// </summary>
		/// <param name="key">The key to add.</param>
		/// <param name="inCaseOfAddition">In case a new key needs to be added, this should be the value.</param>
		/// <param name="a">The action to perform with the key/value.</param>
		public void GetOrAddAndAction(string key, T inCaseOfAddition, Action<T> a)
		{
			if (a == null) return;

			var valueWithLock = allValueWithLock.GetOrAdd(key, (b) => { return new ValueWithLock<T>(inCaseOfAddition, new object()); });

			var myValueLock = valueWithLock.Lock;
			var myValue = valueWithLock.Value;

			lock (myValueLock)
			{
				System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":AddOrGetAndAction entered lock" + key);
				a(myValue);
			}

			System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":AddOrGetAndAction EXIT lock" + key);
		}

		/// <summary>
		/// Tries to get the specific key/value and perform a thread-safe action with it. Will do nothing if key doesn't exist.
		/// </summary>
		/// <param name="key">The key to get.</param>
		/// <param name="a">The action to perform with the key/value.</param>
		public void GetAndAction(string key, Action<T> a)
		{
			if (a == null) return;

			ValueWithLock<T> myValueAndLock;

			if (!allValueWithLock.TryGetValue(key, out myValueAndLock))
			{
				return;
			}

			object myValueLock = myValueAndLock.Lock;
			T myValue = myValueAndLock.Value;

			lock (myValueLock)
			{
				System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":GetAndAction entered lock" + key);
				a(myValue);
			}

			System.Diagnostics.Debug.WriteLine(DateTime.Now.ToLongTimeString() + ":GetAndAction EXIT lock" + key);
		}

		public ICollection<string> GetKeys()
		{
			return allValueWithLock.Keys;
		}

		/// <summary>
		/// Removes the value if it exists. Does nothing if it doesn't exist.
		/// </summary>
		/// <param name="key"></param>
		public void Remove(string key)
		{
			ValueWithLock<T> removedItem;
			allValueWithLock.TryRemove(key, out removedItem);
		}
	}
}