using System;

namespace VChip
{
	class VignereCipher
	{
		public string PlainWord { get; set; } = string.Empty;
		private string Key = string.Empty;
		public string CipherWord { get; set; } = string.Empty;
		public int nextKeyUsedPosition { get; set; } = 0;
	
		public VignereCipher(string key)
        {
			this.Key = key;
        }
			
		public string Encrypt(string word)
        {
			CipherWord = string.Empty;
			int nonAlphaCharCount = 0;
			int keyIndex = 0;
			//Console.WriteLine(word);

			for(int i = 0; i < word.Length; i++)
            {
				if(char.IsLetter(word[i])) {
					bool charIsUpper = char.IsUpper(word[i]);
					char offset = charIsUpper ? 'A' : 'a';
						
					keyIndex = ((nextKeyUsedPosition + i) % Key.Length);
					Console.WriteLine( $"{word[i]}-{i} : {Key[keyIndex]}-{keyIndex}");
					// Logic
					// Mengetahui urutan huruf kunci, A : 0, B :1
					char charKey = (charIsUpper ? char.ToUpper(Key[keyIndex]) : char.ToLower(Key[keyIndex]));

					int numericLetterKey = charToNumeric(charKey, offset);
					int numericLetterPlain = charToNumeric(word[i], offset);
					int numericLetterEncrpyt = crypt(numericLetterPlain,numericLetterKey);
					char charLetterEncrypt = numericToChar(numericLetterEncrpyt, offset);

					// End Logic
						
					CipherWord += charLetterEncrypt;
                } else
                {
					Console.WriteLine($"{word[i]}-{i}");
					CipherWord += word[i];
					nonAlphaCharCount++;
                }
            }

			nextKeyUsedPosition = keyIndex + 1;
			return CipherWord;
        }
	

		public string Decrypt(string word)
		{
			CipherWord = string.Empty;
			int nonAlphaCharCount = 0;
			int keyIndex = 0;
			//Console.WriteLine(word);

			for (int i = 0; i < word.Length; i++)
			{
				if (char.IsLetter(word[i]))
				{
					bool charIsUpper = char.IsUpper(word[i]);
					char offset = charIsUpper ? 'A' : 'a';

					keyIndex = ((nextKeyUsedPosition + i) % Key.Length);
					Console.WriteLine($"{word[i]}-{i} : {Key[keyIndex]}-{keyIndex}");
					// Logic
					// Mengetahui urutan huruf kunci, A : 0, B :1
					char charKey = (charIsUpper ? char.ToUpper(Key[keyIndex]) : char.ToLower(Key[keyIndex]));

					int numericLetterKey = charToNumeric(charKey, offset);
					int numericLetterPlain = charToNumeric(word[i], offset);
					int numericLetterEncrpyt = crypt(numericLetterPlain, numericLetterKey, true);
					char charLetterEncrypt = numericToChar(numericLetterEncrpyt, offset);

					// End Logic

					CipherWord += charLetterEncrypt;
				}
				else
				{
					Console.WriteLine($"{word[i]}-{i}");
					CipherWord += word[i];
					nonAlphaCharCount++;
				}
			}
			nextKeyUsedPosition = keyIndex + 1;
			return CipherWord;
		}

		public void Reset()
        {
			PlainWord = string.Empty;
			Key = string.Empty; 
			CipherWord = string.Empty;
			nextKeyUsedPosition = 0;
		}

		private int charToNumeric(char letter, int offset)
        {
			return (int)letter - offset;
        }

		private char numericToChar(int letter, int offset)
        {
			return (char)((int)letter + offset);
        }

		private int crypt(int letter, int key, bool decrypt = false)
        {
			int result;
			if(decrypt)
            {
				// Decrypt : (Ci - Ki + 26) % 26
				result = ((letter - key) + 26) % 26;
            } else
            {
				// Encrypt : (Pi + Ki) % 26
				result = (letter + key) % 26;
            }

			return result;
        }
	}
}
