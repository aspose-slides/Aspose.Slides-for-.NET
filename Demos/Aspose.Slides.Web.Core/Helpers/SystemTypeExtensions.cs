using System;

namespace Aspose.Slides.Web.Core.Helpers
{
	/// <summary>
	/// The methods-extensions for system types
	/// </summary>
	public static class SystemTypeExtensions
	{
		/// <summary>
		/// Compares two byte arrays
		/// </summary>
		/// <param name="arrayOne"></param>
		/// <param name="arrayTwo"></param>
		/// <returns></returns>
		public static bool Compare(this byte[] arrayOne, byte[] arrayTwo)
		{
			if (arrayOne.Length != arrayTwo.Length)
				return false;

			for (int i = 0; i < arrayOne.Length; i++)
				if (arrayOne[i] != arrayTwo[i])
					return false;

			return true;
		}

		/// <summary>
		/// Parses string as enum of specified type.
		/// </summary>
		/// <typeparam name="T">Enum type.</typeparam>
		/// <param name="source">Source string.</param>
		/// <param name="ignoreCase">Is operation case insensitive.</param>
		/// <returns>Enum value.</returns>
		public static T ParseEnum<T>(this string source, bool ignoreCase = true) where T : Enum => (T)Enum.Parse(typeof(T), source, ignoreCase);
	}
}
