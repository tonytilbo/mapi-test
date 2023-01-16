/*
 * By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * */

// MAPITest.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

using namespace std;

#include <iostream>
#include <fstream>
#include <tchar.h>
#include <locale>
#include <codecvt>
#include <initguid.h>
#define USES_IID_IMAPIFolder
#define USES_IID_IMAPITable
#include <Mapix.h>
#include <mapiutil.h>
#define MDB_ONLINE				((ULONG) 0x00000100)

fstream logFile;
LPMAPISESSION lpSession = NULL;

enum CCSFLAGS
{
	CCSF_SMTP = 0x00000002, // the converter is being passed an SMTP message
	CCSF_NOHEADERS = 0x00000004, // the converter should ignore the headers on the outside message
	CCSF_USE_TNEF = 0x00000010, // the converter should embed TNEF in the MIME message
	CCSF_INCLUDE_BCC = 0x00000020, // the converter should include Bcc recipients
	CCSF_8BITHEADERS = 0x00000040, // the converter should allow 8 bit headers
	CCSF_USE_RTF = 0x00000080, // the converter should do HTML->RTF conversion
	CCSF_PLAIN_TEXT_ONLY = 0x00001000, // the converter should just send plain text
	CCSF_NO_MSGID = 0x00004000, // don't include Message-Id from MAPI message in outgoing messages create a new one
	CCSF_EMBEDDED_MESSAGE = 0x00008000, // We're translating an embedded message
	CCSF_PRESERVE_SOURCE =
	0x00040000, // The convertor should not modify the source message so no conversation index update, no message id, and no header dump.
	CCSF_GLOBAL_MESSAGE = 0x00200000, // The converter should build an international message (EAI/RFC6530)
};

// http://msdn2.microsoft.com/en-us/library/bb905202.aspx
typedef
enum tagENCODINGTYPE
{
	IET_BINARY = 0,
	IET_BASE64 = IET_BINARY + 1,
	IET_UUENCODE = IET_BASE64 + 1,
	IET_QP = IET_UUENCODE + 1,
	IET_7BIT = IET_QP + 1,
	IET_8BIT = IET_7BIT + 1,
	IET_INETCSET = IET_8BIT + 1,
	IET_UNICODE = IET_INETCSET + 1,
	IET_RFC1522 = IET_UNICODE + 1,
	IET_ENCODED = IET_RFC1522 + 1,
	IET_CURRENT = IET_ENCODED + 1,
	IET_UNKNOWN = IET_CURRENT + 1,
	IET_BINHEX40 = IET_UNKNOWN + 1,
	IET_LAST = IET_BINHEX40 + 1
} 	ENCODINGTYPE;
typedef
enum tagMIMESAVETYPE
{
	SAVE_RFC822 = 0,
	SAVE_RFC1521 = SAVE_RFC822 + 1
} 	MIMESAVETYPE;
typedef const struct HCHARSET__* HCHARSET;

typedef HCHARSET* LPHCHARSET;
typedef
enum tagCSETAPPLYTYPE
{
	CSET_APPLY_UNTAGGED = 0,
	CSET_APPLY_ALL = CSET_APPLY_UNTAGGED + 1,
	CSET_APPLY_TAG_ALL = CSET_APPLY_ALL + 1
} 	CSETAPPLYTYPE;

interface IConverterSession : public IUnknown
{
public:
	virtual HRESULT STDMETHODCALLTYPE SetAdrBook(LPADRBOOK pab);

	virtual HRESULT STDMETHODCALLTYPE SetEncoding(ENCODINGTYPE et);

	virtual HRESULT PlaceHolder1();

	virtual HRESULT STDMETHODCALLTYPE MIMEToMAPI(LPSTREAM pstm, LPMESSAGE pmsg, LPCSTR pszSrcSrv, ULONG ulFlags);

	virtual HRESULT STDMETHODCALLTYPE MAPIToMIMEStm(LPMESSAGE pmsg, LPSTREAM pstm, ULONG ulFlags);

	virtual HRESULT PlaceHolder2();
	virtual HRESULT PlaceHolder3();
	virtual HRESULT PlaceHolder4();

	virtual HRESULT STDMETHODCALLTYPE SetTextWrapping(bool fWrapText, ULONG ulWrapWidth);

	virtual HRESULT STDMETHODCALLTYPE SetSaveFormat(MIMESAVETYPE mstSaveFormat);

	virtual HRESULT PlaceHolder5();

	virtual HRESULT STDMETHODCALLTYPE SetCharset(bool fApply, HCHARSET hcharset, CSETAPPLYTYPE csetapplytype);
};

typedef IConverterSession* LPCONVERTERSESSION;

// Class Identifiers
// {4e3a7680-b77a-11d0-9da5-00c04fd65685}
DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

// Interface Identifiers
// {4b401570-b77b-11d0-9da5-00c04fd65685}
DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

void log(string data)
{
	std::cout << data;
	if (logFile)
		logFile << data;
}

void log(string data, LPWSTR lpszW)
{
	string moreData = wstring_convert<codecvt_utf8<wchar_t>, wchar_t>().to_bytes(wstring(lpszW));
	std::cout << data << moreData << "\n";
	if (logFile)
		logFile << data << moreData << "\n";
}


void logError(string data, HRESULT hr)
{
	std::cout << data << hr << "\n";
	if (logFile)
		logFile << data << hr << "\n";
}

void logError(string data, HRESULT hr, LPWSTR lpszW)
{
	string moreData = wstring_convert<codecvt_utf8<wchar_t>, wchar_t>().to_bytes(wstring(lpszW));
	std::cout << data << moreData << "  hr=" << hr << "\n";
	if (logFile)
		logFile << data << moreData << "  hr=" << hr << "\n";
}

enum { EID, EMAIL_ADDRESS, DISPLAY_NAME, DEFAULT_STORE, NUM_COLS };

// Testing the IConverterSession MAPItoMIMEStm
STDMETHODIMP ConvertMessage(LPMESSAGE lpMessage)
{
	LPCONVERTERSESSION lpConverterSession = NULL;
	LPSTREAM lpEMLStm = NULL;
	HRESULT hr;

	if (FAILED(hr = CoCreateInstance(CLSID_IConverterSession, NULL, CLSCTX_INPROC_SERVER, IID_IConverterSession, reinterpret_cast<LPVOID*>(&lpConverterSession))))
	{
		logError("Failed to create IConverterSession", hr);
		return hr;
	}

	if (FAILED(hr = lpConverterSession->SetEncoding(IET_QP)))
	{
		logError("Failed to set encoding type", hr);
		return hr;
	}
	if (FAILED(hr = lpConverterSession->SetSaveFormat(SAVE_RFC1521)))
	{
		logError("Failed to set save format", hr);
		return hr;
	}

	// Create a stream to write the EML to
	if (FAILED(hr = CreateStreamOnHGlobal(NULL, true, &lpEMLStm)))
	{
		logError("Failed to create ouput stream", hr);
		return hr;
	}
	// Call MAPI to MIME conversion
	ULONG ccsFlags = CCSF_SMTP;

	if (FAILED(hr = lpConverterSession->MAPIToMIMEStm(lpMessage, lpEMLStm, ccsFlags)))
	{
		logError("MAPItoMIMEStm call failed", hr);
		return hr;
	}

	return S_OK;
}

// Following adapted from https://learn.microsoft.com/en-us/outlook/troubleshoot/development/how-to-list-messages-in-inbox-with-mapi

/// <summary>
/// List the subject of each item found in the given folder
/// </summary>
/// <param name="lpMDB">Message store</param>
/// <param name="lpMessageFolder">Folder</param>
/// <returns>HRESULT</returns>
STDMETHODIMP ListMessages(
	LPMDB lpMDB,
	LPMAPIFOLDER lpMessageFolder)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpContentsTable = NULL;
	LPSRowSet pRows = NULL;
	LPSTREAM lpStream = NULL;
	ULONG i;

	// Define a SPropTagArray array here using the SizedSPropTagArray Macro
	// This enum will allows you to access portions of the array by a name instead of a number.
	// If more tags are added to the array, appropriate constants need to be added to the enum.
	enum {
		ePR_SUBJECT,
		ePR_ENTRYID,
		MSG_NUM_COLS
	};
	// These tags represent the message information we want to retrieve
	static SizedSPropTagArray(MSG_NUM_COLS, sptCols) = { MSG_NUM_COLS,
		PR_SUBJECT,
		PR_ENTRYID
	};

	log("Attempting to list items in folder\n");
	hRes = lpMessageFolder->GetContentsTable(
		0,
		&lpContentsTable);
	if (FAILED(hRes))
	{
		logError("Failed on GetContentsTable: ", hRes);
		goto quit;
	}

	hRes = HrQueryAllRows(
		lpContentsTable,
		(LPSPropTagArray)&sptCols,
		NULL, // restriction...we're not using this parameter
		NULL, // sort order...we're not using this parameter
		0,
		&pRows);
	if (FAILED(hRes))
	{
		logError("Failed on HrQueryAllRows: ", hRes);
		goto quit;
	}

	// Get the first message only
	for (i = 0; i < 1; i++)
	{
		LPMESSAGE lpMessage = NULL;
		ULONG ulObjType = NULL;
		LPSPropValue lpProp = NULL;

		if (PR_SUBJECT == pRows->aRow[i].lpProps[ePR_SUBJECT].ulPropTag)
			log("Item subject: ", pRows->aRow[i].lpProps[ePR_SUBJECT].Value.lpszW);

		hRes = lpMDB->OpenEntry(
			pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
			(LPENTRYID)pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb,
			NULL, // default interface
			MAPI_BEST_ACCESS,
			&ulObjType,
			(LPUNKNOWN*)&lpMessage);

		if (!FAILED(hRes))
		{
			// We've opened the message
			// We don't do anything further at this point, but the message can be accessed using lpMessage
			ConvertMessage(lpMessage);
		}
		else
			logError("OpenEntry error: ", hRes);

		MAPIFreeBuffer(lpProp);
		UlRelease(lpMessage);
		hRes = S_OK;
	}

quit:
	FreeProws(pRows);
	UlRelease(lpContentsTable);
	return hRes;
}

/// <summary>
/// Attempt to open the receive folder for the given message store
/// </summary>
/// <param name="lpMDB">The message store</param>
/// <param name="lpInboxFolder">Receive folder (if found)</param>
/// <returns>HRESULT</returns>
STDMETHODIMP OpenInbox(
	LPMDB lpMDB,
	LPMAPIFOLDER* lpInboxFolder)
{
	ULONG cbInbox;
	LPENTRYID lpbInbox;
	ULONG ulObjType;
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpTempFolder = NULL;
	LPSPropValue tmp = NULL;

	*lpInboxFolder = NULL;

	hRes = lpMDB->GetReceiveFolder(
		NULL, // Get default receive folder
		NULL, // Flags
		&cbInbox,
		&lpbInbox,
		NULL);
	if (FAILED(hRes))
	{
		logError("Failed on GetReceiveFolder: ", hRes);
		goto quit;
	}

	hRes = lpMDB->OpenEntry(
		cbInbox, // Size and...
		lpbInbox, // Value of the Inbox's EntryID
		NULL, // We want the default interface (IMAPIFolder)
		MAPI_BEST_ACCESS, // Flags
		&ulObjType, // Object returned type
		(LPUNKNOWN*)&lpTempFolder); //Returned folder
	if (FAILED(hRes))
	{
		logError("Failed on OpenEntry (receive folder): ", hRes);
		goto quit;
	}

	// Retrieve and log the name of the folder we have opened
	hRes = HrGetOneProp(
		lpTempFolder,
		PR_DISPLAY_NAME,
		&tmp);
	if (FAILED(hRes)) goto quit;
	log("Opened receive folder: ", tmp->Value.lpszW);

	// Assign the out parameter
	*lpInboxFolder = lpTempFolder;

quit:
	if (tmp) MAPIFreeBuffer(tmp);
	MAPIFreeBuffer(lpbInbox);
	return hRes;
}

/// <summary>
/// Process the specified message store (attempts to open store and then list messages in default receive folder)
/// </summary>
/// <param name="storeInfoRow">SRow from stores table containing message store info (e.g. EntryId)</param>
/// <returns>HRESULT</returns>
HRESULT ProcessMessageStore(SRow storeInfoRow)
{
	LPMDB       pMDB = NULL;

	HRESULT hr = lpSession->OpenMsgStore(NULL,
		storeInfoRow.lpProps[EID].Value.bin.cb,
		(LPENTRYID)storeInfoRow.lpProps[EID].Value.bin.lpb,
		NULL,
		MAPI_BEST_ACCESS | MDB_NO_DIALOG | MDB_ONLINE,
		&pMDB);

	if (SUCCEEDED(hr))
	{
		// We have a message store, so try to open the root folder
		log("Opened message store: ", storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);

		LPUNKNOWN lpUnk = NULL;
		ULONG ulType = 0;
		hr = pMDB->OpenEntry(0, NULL, NULL, MAPI_BEST_ACCESS, &ulType, &lpUnk);
		if (SUCCEEDED(hr))
		{
			LPMAPIFOLDER msgStoreRoot = (LPMAPIFOLDER)lpUnk;
			log("Opened root folder of message store: ", storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);

			// Open inbox
			LPMAPIFOLDER inbox;
			hr = OpenInbox(pMDB, &inbox);
			if (SUCCEEDED(hr))
			{
				ListMessages(pMDB, inbox);
				// Release inbox
				inbox->Release();
				log("Receive folder released\n");
			}

			msgStoreRoot->Release();
			log("Released root folder of message store: ", storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);
		}
		else
			logError("Failed to open root folder of message store: ", hr, storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);

		pMDB->Release();
		log("Released message store: ", storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);
	}
	else
		logError("Failed on OpenMsgStore: ", hr, storeInfoRow.lpProps[DISPLAY_NAME].Value.lpszW);

	return hr;
}

/// <summary>
/// Initialise MAPI, log on and read the message store table
/// </summary>
/// <returns>0 if successful, error code otherwise</returns>
int MAPITest()
{
	int nRetCode = 0;
	HRESULT hRes = 0, hr=0;
	LPWSTR lpszProfile = NULL;

	log("Initialising MAPI test\n");

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL)
	{
		// Create log file
		log("Creating log file\n");
		logFile.open("MAPITest Log.txt", ios::out);
		if (!logFile)
			log("Failed to create log file\n");

		log("Starting MAPI log-on test\n");

		// Initialize MAPI
		if FAILED(hRes = MAPIInitialize(0))
		{
			logError("Fatal Error: MAPIInitialize failed ", hRes);
			nRetCode = hRes;
		}
		else
		{
			log("MAPIInitialize succeeded\n");
			FLAGS flags = MAPI_NEW_SESSION | MAPI_EXTENDED | MAPI_USE_DEFAULT;

			// Log on to MAPI session
			if (FAILED(hRes = MAPILogonEx(NULL, lpszProfile, NULL, flags, &lpSession)))
				logError("Error: MAPILogonEx failed ", hRes);
			else
			{
				log("MAPILogonEx succeeded\n");
				// Retrieve list of message stores
				LPMAPITABLE pStoresTbl = NULL;
				SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS, PR_ENTRYID, PR_EMAIL_ADDRESS, PR_DISPLAY_NAME, PR_DEFAULT_STORE };

				HRESULT hr = lpSession->GetMsgStoresTable(0, &pStoresTbl);
				if (FAILED(hr))
				{
					nRetCode = hr;
					logError("GetMsgStoresTable failed: ", hr);
				}
				else
				{
					log("GetMsgStoresTable succeeded\n");

					LPSRowSet   rows = NULL;
					LPMDB       pMDB = NULL;
					hr = HrQueryAllRows(pStoresTbl, (LPSPropTagArray)&sptCols, NULL, NULL, 0, &rows);
					if (SUCCEEDED(hr))
					{
						// We have a list of message stores... Display them, and attempt to open the default store
						for (unsigned int i = 0; i < rows->cRows; i++)
						{
							if (rows->aRow[i].lpProps[DEFAULT_STORE].Value.b)
							{
								// This is the default store
								log("Default message store: ", rows->aRow[i].lpProps[DISPLAY_NAME].Value.lpszW);
								ProcessMessageStore(rows->aRow[i]);
							}
							else
								log("Message store: ", rows->aRow[i].lpProps[DISPLAY_NAME].Value.lpszW);
						}
					}
					if (rows) FreeProws(rows);
					pStoresTbl->Release();
				}

				// Log off
				if (FAILED(hRes = lpSession->Logoff(0, 0, 0)))
					logError("Error at Logoff: ", hRes);
				else
					log("LogOff succeeded\n");
			}
			if (lpSession) lpSession->Release();
			MAPIUninitialize();
		}
		logFile.close();
	}
	return nRetCode;
}

/// <summary>
/// Main entry point
/// </summary>
/// <returns>0 if successful, error code otherwise</returns>
int main()
{
	return MAPITest();
}



