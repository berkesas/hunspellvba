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
#pragma once

#include <Windows.h>
#include <oleauto.h>
#include "../../hunspell-1.7.2/src/hunspell/hunspell.hxx"

extern "C" {
	__declspec(dllexport) void __stdcall HunspellInit(Hunspell** hunspell, const char* affixFilePath, const char* dictionaryFilePath);
	__declspec(dllexport) bool __stdcall CheckSpelling(Hunspell* hunspell, BSTR word);
	__declspec(dllexport) void __stdcall HunspellFree(Hunspell* hunspell);
	__declspec(dllexport) int __stdcall AddDictionary(Hunspell* hunspell, const char* dictionaryFilePath);

	/**
	 * @brief Suggest words based on combination of affix+roots.
	 *
	 * This function uses Hunspell handle, bad string, and a pointer
	 * It tries various small edits (like removing, inserting, or swapping letters) 
	 * to generate possible corrections. If no good suggestions are found, 
	 * it then tries ngram-based suggestions (calculating ngram-based distance to all dictionary words) 
	 * and phonetic-based suggestions
	 *
	 * @param hunspell - handle to Hunspell created by HunspellInit
	 * @param word - misspelling for which suggestions should be generated
	 * @param count - pointer to the number of suggestions returned 
	 * @return pointer to UTF-8 string array
	 *
	 * @pre hunspell, count must be created before.
	 * @post returned pointer must be freed by the calling side.
	 */
	__declspec(dllexport) const char** __stdcall GetSuggestions(Hunspell* hunspell, BSTR word, int* count);

	/**
	 * @brief Suggest words based on combination of affix+roots, more focused than suggest()
	 *
	 * This function The suffix_suggest function specifically generates 
	 * suggestions based on suffixes. It calculates the similarity of the 
	 * misspelled word to dictionary word stems and produces suggestions 
	 * by applying possible suffixes. It focuses on suffixes and is typically used 
	 * when you want to generate suggestions that involve adding or modifying suffixes.
	 *
	 * @param hunspell - handle to Hunspell created by HunspellInit
	 * @param word - misspelling for which suggestions should be generated
	 * @param count - pointer to the number of suggestions returned
	 * @return pointer to UTF-8 string array
	 *
	 * @pre hunspell, count must be created before.
	 * @post returned pointer must be freed by the calling side.
	 */
	__declspec(dllexport) const char** __stdcall GetSuffixSuggestions(Hunspell* hunspell, BSTR word, int* count);
	__declspec(dllexport) const char** __stdcall GetMisspellings(Hunspell* hunspell, BSTR text, int* count);
	__declspec(dllexport) void __stdcall FreeItems(const char** items, int count);
	__declspec(dllexport) int __stdcall AddWord(Hunspell* hunspell, BSTR word);
}
