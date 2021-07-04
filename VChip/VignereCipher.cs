namespace VChip
{
	class VignereCipher
	{
		private static int Mod(int a, int b)
		{
			return ((a % b) + b) % b;
		}

		private static string Cipher(string input, string key, bool encipher)
		{
			// kunci harus huruflp
			for (int i = 0; i < key.Length; ++i)
				if (!char.IsLetter(key[i]))
					return null; // Error

			string output = string.Empty;

			// total bukan huruf
			int nonAlphaCharCount = 0;

			// Untuk setiap huruf dalam teks
			for (int i = 0; i < input.Length; ++i)
			{
				// Jika itu huruf lakukan ini
				if (char.IsLetter(input[i]))
				{
					// PROBLEM : b (5)
					// KEY = p (0)
					bool cIsUpper = char.IsUpper(input[i]);
					char offset = cIsUpper ? 'A' : 'a'; // A=65 a=97
					int keyIndex = (i - nonAlphaCharCount) % key.Length; // 0
					int k = (cIsUpper ? char.ToUpper(key[keyIndex]) : char.ToLower(key[keyIndex])) - offset;
					k = encipher ? k : -k; // -	15
					char ch = (char)((Mod(((input[i] + k) - offset), 26)) + offset);
					output += ch;
				}
				else
				{
					output += input[i];
					++nonAlphaCharCount;
				}
			}

			return output;
		}

		public static string Encipher(string input, string key)
		{
			return Cipher(input, key, true);
		}

		public static string Decipher(string input, string key)
		{
			return Cipher(input, key, false);
		}
	}
}
