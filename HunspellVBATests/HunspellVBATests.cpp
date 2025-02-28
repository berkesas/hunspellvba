/*
 * MIT License
 *
 * Copyright (c) 2024-2025 Nazar Mammedov
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
#include "pch.h"
#include <iostream>
#include "CppUnitTest.h"
#include "../HunspellVBA/HunspellVBA.h"
#include "../HunspellVBA/HunspellVBA.cpp"
#include <Windows.h>
#include <oleauto.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace HunspellVBATests
{
	TEST_CLASS(HunspellVBATests)
	{
	public:

		TEST_METHOD(HunspellInitTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";

			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);

			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			delete hunspell;
		}

		TEST_METHOD(HunspellInitInvalidParametersTest)
		{
			const char* validAffixFilePath = "lang/tk-TM.aff";
			const char* validDictionaryFilePath = "lang/tk-TM.dic";

			Hunspell* hunspell = nullptr;

			// Test 1: Null pointer for Hunspell** parameter
			try {
				HunspellInit(nullptr, validAffixFilePath, validDictionaryFilePath);
				Assert::Fail(L"Expected exception not thrown for nullptr Hunspell** parameter.");
			}
			catch (const std::invalid_argument& e) {
				Assert::AreEqual("Null pointer argument.", e.what());
			}
			catch (...) {
				Assert::Fail(L"Unexpected exception type thrown.");
			}

			// Test 2: Null pointer for affixFilePath parameter
			try {
				HunspellInit(&hunspell, nullptr, validDictionaryFilePath);
				Assert::Fail(L"Expected exception not thrown for nullptr affixFilePath parameter.");
			}
			catch (const std::invalid_argument& e) {
				Assert::AreEqual("Null pointer argument.", e.what());
			}
			catch (...) {
				Assert::Fail(L"Unexpected exception type thrown.");
			}

			// Test 3: Null pointer for dictionaryFilePath parameter
			try {
				HunspellInit(&hunspell, validAffixFilePath, nullptr);
				Assert::Fail(L"Expected exception not thrown for nullptr dictionaryFilePath parameter.");
			}
			catch (const std::invalid_argument& e) {
				Assert::AreEqual("Null pointer argument.", e.what());
			}
			catch (...) {
				Assert::Fail(L"Unexpected exception type thrown.");
			}

			delete hunspell;
			hunspell = nullptr;
		}



		TEST_METHOD(CheckSpellingTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			const wchar_t* source = L"şagal";
			BSTR word = SysAllocString(source);
			bool result = CheckSpelling(hunspell, word);
			Assert::IsTrue(result, L"4 The word 'şagal' should be spelled correctly");

			source = L"zanar";
			word = SysAllocString(source);
			result = CheckSpelling(hunspell, word);
			Assert::IsFalse(result, L"The word 'zanar' should be misspelled");

			HunspellFree(hunspell);
		}

		TEST_METHOD(CheckSpellingTextTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			int count;
			const wchar_t* source = L"Adamlar bazarak gitdiler. Hemme zatd gowy bolarmyka?";
			BSTR word = SysAllocString(source);
			const char** items = GetMisspellings(hunspell, word, &count);

			Assert::AreNotEqual(0, count, L"No suggestions returned");

			std::ofstream logFile2("spellcheck_log.txt", std::ios::app);
			if (logFile2.is_open()) {
				for (int i = 0; i < count; ++i) {
					std::cout << items[i] << std::endl;
					logFile2 << items[i] << std::endl;
				}
				logFile2.flush();
				logFile2.close();
			}

			FreeItems(items, count);

			HunspellFree(hunspell);
		}

		TEST_METHOD(GetSuggestionsTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			int count;
			const wchar_t* source = L"bazal";
			BSTR word = SysAllocString(source);
			const char** suggestions = GetSuggestions(hunspell, word, &count);

			Assert::AreNotEqual(0, count, L"No suggestions returned");

			std::cout << "Suggestions" << std::endl;
			for (int i = 0; i < count; ++i) {
				std::cout << suggestions[i] << std::endl;
			}

			FreeItems(suggestions, count);

			HunspellFree(hunspell);
		}

		TEST_METHOD(GetSuffixSuggestionsTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			int count;
			const wchar_t* source = L"akla";
			BSTR word = SysAllocString(source);
			const char** suggestions = GetSuffixSuggestions(hunspell, word, &count);

			Assert::AreNotEqual(0, count, L"No suggestions returned");

			std::cout << "Suffix suggestions" << std::endl;
			if (count > 10) count = 10;
			for (int i = 0; i < count; ++i) {
				std::cout << suggestions[i] << std::endl;
			}

			FreeItems(suggestions, count);

			HunspellFree(hunspell);
		}

		TEST_METHOD(AddDictionaryTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			const wchar_t* source = L"gerontologiýa";
			BSTR word = SysAllocString(source);
			bool result = CheckSpelling(hunspell, word);
			Assert::IsFalse(result, L"The word 'gerontologiýa' should be misspelled");

			const char* additionalDictionaryFilePath = "lang/tk-TM_addition.dic";

			result = AddDictionary(hunspell, additionalDictionaryFilePath);

			Assert::IsTrue((result == 0), L"Dictionary is loaded");

			source = L"gerontologiýa";
			word = SysAllocString(source);
			result = CheckSpelling(hunspell, word);
			Assert::IsTrue(result, L"The word 'gerontologiýa' should be spelled correctly");

			HunspellFree(hunspell);
		}

		TEST_METHOD(AddWordTest)
		{
			const char* affixFilePath = "lang/tk-TM.aff";
			const char* dictionaryFilePath = "lang/tk-TM.dic";
			Hunspell* hunspell = nullptr;

			HunspellInit(&hunspell, affixFilePath, dictionaryFilePath);
			Assert::IsNotNull(hunspell, L"Failed to initialize Hunspell");

			const wchar_t* source = L"Bolmajaksöz";
			BSTR word = SysAllocString(source);
			bool result = CheckSpelling(hunspell, word);
			Assert::IsFalse(result, L"The word 'Bolmajaksöz' should be misspelled");

			source = L"Bolmajaksöz";
			word = SysAllocString(source);
			result = AddWord(hunspell, word);
			Assert::IsTrue((result == 0), L"Word is added");

			source = L"Bolmajaksöz";
			word = SysAllocString(source);
			result = CheckSpelling(hunspell, word);
			Assert::IsTrue(result, L"The word 'Bolmajaksöz' should be spelled correctly");

			HunspellFree(hunspell);
		}
	};
}