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
#include "HunspellVBA.h"
#include <string>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <vector>
#include <Windows.h>
#include <sstream>
#include "utf8.h"
#include <stdexcept>

void __stdcall HunspellInit(Hunspell** hunspell, const char* affixFilePath, const char* dictionaryFilePath) {
	if (hunspell == nullptr || affixFilePath == nullptr || dictionaryFilePath == nullptr) {
#ifdef _DEBUG
		std::cerr << "Error: Null pointer argument." << std::endl;
#endif
		throw std::invalid_argument("Null pointer argument.");
	}

	try {
		*hunspell = new Hunspell(affixFilePath, dictionaryFilePath);
	}
	catch (const std::exception& e) {
#ifdef _DEBUG
		std::cerr << "Hunspell initialization failed: " << e.what() << std::endl;
#endif
		* hunspell = nullptr;
	}
}

bool __stdcall CheckSpelling(Hunspell* hunspell, BSTR word) {
	if (hunspell == nullptr) {
		std::cerr << "Error: Null pointer passed for Hunspell instance." << std::endl;
		return false;
	}

	if (word == nullptr) {
		std::cerr << "Error: Null pointer passed for word." << std::endl;
		return false;
	}

	try {
		std::wstring wstr(word, SysStringLen(word));

		int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), NULL, 0, NULL, NULL);
		if (size_needed <= 0) {
			std::cerr << "Error: WideCharToMultiByte failed to calculate size." << std::endl;
			return false;
		}

		std::vector<char> utf8_buffer(size_needed);
		int result = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), utf8_buffer.data(), size_needed, NULL, NULL);
		if (result == 0) {
			std::cerr << "Error: WideCharToMultiByte conversion failed." << std::endl;
			return false;
		}

		std::string utf8str(utf8_buffer.begin(), utf8_buffer.end());

		return hunspell->spell(utf8str);
	}
	catch (const std::exception& ex) {
		std::cerr << "Exception: " << ex.what() << std::endl;
		return false;
	}
	catch (...) {
		std::cerr << "Unknown error occurred during CheckSpelling." << std::endl;
		return false;
	}
}



void __stdcall HunspellFree(Hunspell* hunspell) {
	if (hunspell != nullptr) {
		delete hunspell;
		hunspell = nullptr;
	}
}

int __stdcall AddDictionary(Hunspell* hunspell, const char* dictionaryFilePath) {
	if (hunspell == nullptr) {
		std::cerr << "Error: Null pointer passed for Hunspell instance." << std::endl;
		return -1;
	}

	if (dictionaryFilePath == nullptr) {
		std::cerr << "Error: Null pointer passed for dictionary file path." << std::endl;
		return -2;
	}

	try {
		int result = hunspell->add_dic(dictionaryFilePath);

		if (result != 0) {
			std::cerr << "Error: Failed to add dictionary. Result code: " << result << std::endl;
			return result;
		}

		return 0;
	}
	catch (const std::exception& ex) {
		std::cerr << "Exception: " << ex.what() << std::endl;
		return -3;
	}
	catch (...) {
		std::cerr << "Unknown error occurred while adding dictionary." << std::endl;
		return -4;
	}
}

int __stdcall AddWord(Hunspell* hunspell, BSTR word) {
	if (hunspell == nullptr) {
		std::cerr << "Error: Null pointer passed for Hunspell instance." << std::endl;
		return -1;
	}

	if (word == nullptr) {
		std::cerr << "Error: Null pointer passed for word." << std::endl;
		return -2;
	}

	try {
		std::wstring wstr(word, SysStringLen(word));

		int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), NULL, 0, NULL, NULL);
		if (size_needed <= 0) {
			std::cerr << "Error: WideCharToMultiByte failed to calculate size." << std::endl;
			return -3;
		}

		std::vector<char> utf8_buffer(size_needed);
		int result = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), utf8_buffer.data(), size_needed, NULL, NULL);
		if (result == 0) {
			std::cerr << "Error: WideCharToMultiByte conversion failed." << std::endl;
			return -4;
		}

		std::string utf8str(utf8_buffer.begin(), utf8_buffer.end());

		return hunspell->add(utf8str);
	}
	catch (const std::exception& ex) {
		std::cerr << "Exception: " << ex.what() << std::endl;
		return -5;
	}
	catch (...) {
		std::cerr << "Unknown error occurred while adding word." << std::endl;
		return -6;
	}
}


void __stdcall FreeItems(const char** items, int count) {
	for (int i = 0; i < count; ++i) {
		free((void*)items[i]);
	}
	free(items);
}

const char** __stdcall GetSuggestions(Hunspell* hunspell, BSTR word, int* count) {
	if (count == nullptr) {
		return nullptr;
	}

	std::wstring wstr(word, SysStringLen(word));

	int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), NULL, 0, NULL, NULL);
	std::string utf8str(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), &utf8str[0], size_needed, NULL, NULL);

	std::vector<std::string> suggestions = hunspell->suggest(utf8str);

	const char** result = (const char**)malloc(suggestions.size() * sizeof(const char*));
	for (size_t i = 0; i < suggestions.size(); ++i) {
		result[i] = _strdup(suggestions[i].c_str());
	}

	*count = static_cast<int>(suggestions.size());

	return result;
}

const char** __stdcall GetSuffixSuggestions(Hunspell* hunspell, BSTR word, int* count) {
	if (count == nullptr) {
		return nullptr;
	}

	std::wstring wstr(word, SysStringLen(word));

	int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), NULL, 0, NULL, NULL);
	std::string utf8str(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), &utf8str[0], size_needed, NULL, NULL);

	std::vector<std::string> suggestions = hunspell->suffix_suggest(utf8str);

	const char** result = (const char**)malloc(suggestions.size() * sizeof(const char*));
	for (size_t i = 0; i < suggestions.size(); ++i) {
		result[i] = _strdup(suggestions[i].c_str());
	}

	*count = static_cast<int>(suggestions.size());

	return result;
}

const char** __stdcall GetMisspellings(Hunspell* hunspell, BSTR text, int* count) {
	if (count == nullptr) {
		return nullptr;
	}

	std::wstring wstr(text, SysStringLen(text));

	int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), NULL, 0, NULL, NULL);
	std::string utf8str(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.length(), &utf8str[0], size_needed, NULL, NULL);

	std::vector<std::string> misspelledWords;
	std::string word;
	std::istringstream iss(utf8str);

	while (iss >> word) {
		if (!hunspell->spell(word)) {
			misspelledWords.push_back(word);
		}
	}

	const char** result = (const char**)malloc((misspelledWords.size() + 1) * sizeof(const char*));

	for (size_t i = 0; i < misspelledWords.size(); ++i) {
		result[i] = _strdup(misspelledWords[i].c_str());
	}

	if (result)
	{
		result[misspelledWords.size()] = nullptr;
	}

	if (count)
	{
		*count = static_cast<int>(misspelledWords.size());
	}

	return result;
}