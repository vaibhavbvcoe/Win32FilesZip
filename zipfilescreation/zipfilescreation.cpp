// zipfilescreation.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <comdef.h>
#include <comutil.h>
#pragma comment( lib, "comsuppw.lib" )
#include <Shldisp.h>
#include <tlhelp32.h>

#include <algorithm>
#include <iterator>
#include <memory>
#include <set>
#include <vector>


std::set<DWORD> getAllThreadIds()
{
	auto processId = GetCurrentProcessId();
	auto currThreadId = GetCurrentThreadId();
	std::set<DWORD> thread_ids;
	std::unique_ptr< void, decltype(&CloseHandle) > h(CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0), CloseHandle);
	if (h.get() != INVALID_HANDLE_VALUE)
	{
		THREADENTRY32 te;
		te.dwSize = sizeof(te);
		if (Thread32First(h.get(), &te))
		{
			do
			{
				if (te.dwSize >= (FIELD_OFFSET(THREADENTRY32, th32OwnerProcessID) + sizeof(te.th32OwnerProcessID)))
				{
					//only enumerate threads that are called by this process and not the main thread
					if ((te.th32OwnerProcessID == processId) && (te.th32ThreadID != currThreadId))
					{
						thread_ids.insert(te.th32ThreadID);
					}
				}
				te.dwSize = sizeof(te);
			} while (Thread32Next(h.get(), &te));
		}
	}
	return thread_ids;
}


HRESULT CopyItems(PCWSTR srcDir, PCWSTR destZipDir)
{
	FILE* f;
	_wfopen_s(&f, destZipDir, L"wb");
	// see Data in HKEY_CLASSES_ROOT\.zip\CompressedFolder\ShellNew
	// same value here
	fwrite("\x50\x4B\x05\x06\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00", 22, 1, f);
	fclose(f);


	_COM_SMARTPTR_TYPEDEF(IShellDispatch, IID_IShellDispatch);
	_COM_SMARTPTR_TYPEDEF(Folder, IID_Folder);
	IShellDispatchPtr shell;
	FolderPtr destFolder, srcFolder;

	variant_t dirName, fileName, options;
	CoInitialize(NULL);
	HRESULT hr = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, IID_IShellDispatch, (void**)&shell);
	if (SUCCEEDED(hr))
	{
		dirName = destZipDir;
		hr = shell->NameSpace(dirName, &destFolder);
		if (SUCCEEDED(hr))
		{
			auto existingThreadIds = getAllThreadIds();
			options = FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI;  //NOTE:  same result as 0x0000

			//vaibhav [Start]
			// 
			variant_t vtSrcDir = srcDir;

			shell->NameSpace(vtSrcDir, &srcFolder);

			FolderItems* fi = NULL;
			srcFolder->Items(&fi);

			long count;
			fi->get_Count(&count);

			for (int i = 0; i < count; i++)
			{
				variant_t vFolderItmIdx((long)i, VT_I4);

				FolderItem* item = NULL;

				fi->Item(vFolderItmIdx, &item);

				BSTR itemName;
				item->get_Path(&itemName);

				fileName = itemName;
				hr = destFolder->CopyHere(fileName, options); //NOTE: this appears to always return S_OK even on error

				auto updatedThreadIds = getAllThreadIds();
				std::vector<decltype(updatedThreadIds)::value_type> newThreadIds;
				std::set_difference(updatedThreadIds.begin(), updatedThreadIds.end(), existingThreadIds.begin(), existingThreadIds.end(), std::back_inserter(newThreadIds));

				std::vector<HANDLE> threads;
				for (auto threadId : newThreadIds)
					threads.push_back(OpenThread(SYNCHRONIZE, FALSE, threadId));

				if (!threads.empty())
				{
					// Waiting for new threads to finish not more than 5 min.
					WaitForMultipleObjects(threads.size(), &threads[0], TRUE, 5 * 60 * 1000);

					for (size_t i = 0; i < threads.size(); i++)
						CloseHandle(threads[i]);
				}
			}
		}
	}
	CoUninitialize();
	return hr;
}

int main()
{
	CopyItems(L"F:\\SampleImg",L"F:\\ZipCreated.zip");
	return 0;
}
