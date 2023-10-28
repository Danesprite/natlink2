/*
 Python Macro Language for Dragon NaturallySpeaking
	(c) Copyright 1999 by Joel Gould
	Portions (c) Copyright 1999 by Dragon Systems, Inc.

 appsupp.cpp
	This module implements the COM interface which Dragon NaturallySpeaking
	calls when it connects with a compatibility module.  This implementation
	is designed to be a global client and not a app-specific client.  That
	decision simplifies the design somewhat.
*/
#include <string>
#include "../StdAfx.h"
#include "../Resource.h"
#include "../DragonCode.h"
#include "appsupp.h"
// #include <plog/Log.h>

// from PythWrap.cpp
CDragonCode * initModule();
PCCHAR parsePyString( PyObject * pyWord, const char * encoding );
PCCHAR * parseStringArray( const char * funcName, PyObject * args );
PCCHAR parsePyErrString();
PyObject * executePyCodeAsModule( PCCHAR codeString, char * moduleName );

/////////////////////////////////////////////////////////////////////////////
// CDgnAppSupport

//---------------------------------------------------------------------------

CDgnAppSupport::CDgnAppSupport()
{
	m_pNatLinkMain = NULL;
	m_pDragCode = NULL;
}

//---------------------------------------------------------------------------

CDgnAppSupport::~CDgnAppSupport()
{
}


//---------------------------------------------------------------------------
// Called by NatSpeak once when the compatibility module is first loaded.
// This will never be called more than once in normal use.  (Although if one
// compatibility module calls another as is occassionally the case it could
// be called more than once.  Needless to say that does not apply for this
// project.)
//
// NatSpeak passes in a site object which saves us the trouble of finding
// one ourselves.  NatSpeak will be running.

STDMETHODIMP CDgnAppSupport::Register( IServiceProvider * pIDgnSite )
{
	BOOL bSuccess;
	PCCHAR errorMessage;

	// load and initialize the Python system
	Py_Initialize();

	// load the natlink module into Python and return a pointer to the
	// shared CDragonCode object
	m_pDragCode = initModule();
	m_pDragCode->setAppClass( this );


	// simulate calling natlink.natConnect() except share the site object
	bSuccess = m_pDragCode->natConnect( pIDgnSite );


	if( !bSuccess )
	{
		OutputDebugString(
			TEXT( "NatLink: failed to initialize NatSpeak interfaces") ); // RW TEXT macro added
		m_pDragCode->displayText(
			"Failed to initialize NatSpeak interfaces\r\n", TRUE );
		return S_OK;
	}

	// attempt to add the natlink core directory to sys.path.
	addCoreToSysPath();

	// now load the Python code which sets all the callback functions
	m_pDragCode->setDuringInit( TRUE );
	m_pNatLinkMain = PyImport_ImportModule( "natlinkmain" );
	m_pDragCode->setDuringInit( FALSE );


	if( m_pNatLinkMain == NULL ) {
		OutputDebugString(
			TEXT( "NatLink: an exception occurred loading 'natlinkmain' module" ) ); // RW TEXT macro added
		m_pDragCode->displayText(
			"An exception occurred loading 'natlinkmain' module:\r\n", TRUE );
		
		errorMessage = parsePyErrString();
		if( errorMessage )
		{
			m_pDragCode->displayText( errorMessage, TRUE );
			m_pDragCode->displayText( "\r\n" );
		}
	}

	return S_OK;
}

//---------------------------------------------------------------------------
// Called by NatSpeak during shutdown as the last call into this
// compatibility module.  There is always one UnRegister call for every
// Register call (all one of them).

STDMETHODIMP CDgnAppSupport::UnRegister()
{
	// simulate calling natlink.natDisconnect()
	m_pDragCode->natDisconnect();

	// free our reference to the Python modules
	Py_XDECREF( m_pNatLinkMain );

	return S_OK;
}

//---------------------------------------------------------------------------
// For a non-global client, this call is made evrey time a new instance of
// the target application is started.  The process ID of the target
// application is passed in along with the target application module name
// and a string which tells us where to find NatSpeak information specific
// to that module in the registry.
//
// For global clients (like us), this is called once after Register and we
// can ignore the call.

#ifdef UNICODE
	STDMETHODIMP CDgnAppSupport::AddProcess(
		DWORD dwProcessID,
		const wchar_t * pszModuleName,
		const wchar_t * pszRegistryKey,
		DWORD lcid )
	{
		return S_OK;
	}
#else
	STDMETHODIMP CDgnAppSupport::AddProcess(
		DWORD dwProcessID,
		const char * pszModuleName,
		const char * pszRegistryKey,
		DWORD lcid )
	{
		return S_OK;
	}
#endif
//---------------------------------------------------------------------------
// For a non-global client, this call is made whenever the application whose
// process ID was passed to AddProcess terminates.
//
// For global clients (like us), this is called once just before UnRegister
// and we can ignore the call.

STDMETHODIMP CDgnAppSupport::EndProcess( DWORD dwProcessID )
{
	return S_OK;
}

//---------------------------------------------------------------------------
// This utility function reloads the Python interpreter.  It is called from
// the display window menu and is useful for debugging during development of
// natlinkmain and natlinkutils. In normal use, we do not need to reload the
// Python interpreter.

void CDgnAppSupport::reloadPython()
{
	PyImport_ReloadModule(m_pNatLinkMain);
}

//---------------------------------------------------------------------------
// This finds and adds the Natlink "core" directory to sys.path, if possible.
// Called in Register().

BOOL CDgnAppSupport::addCoreToSysPath()
{
	char * moduleName;
	const char * codeString, * errorMessage;
	PyObject * pModule, * pSysModules;

	/*
	* https://www.python.org/dev/peps/pep-0514/
	* According to PEP514 python should scan this registry location when
	* it builds sys.path when the interpreter is initialised. At least on
	* my system this is not happening correctly, and natlinkmain is not being
	* found. This code pulls the value (set by the config scripts) from the
	* registry manually and adds it to the module search path.
	*
	* Exceptions raised here will not cause a crash, so the worst case scenario
	* is that we add a value to the path which is already there.
	*/

	// NOTE: Ensure the following Python code works in each supported Python
	//       version; the stable ABI does not help us here.

	moduleName = "TempPyModule";
	codeString =
		"import winreg, sys, traceback\n"
		"hive = winreg.HKEY_LOCAL_MACHINE\n"
		"key = \"Software\\\\Python\\\\PythonCore\\\\\" + sys.winver + \"\\\\PythonPath\\\\Natlink\"\n"
		"flags = winreg.KEY_READ | winreg.KEY_WOW64_32KEY\n"
		"natlink_key = winreg.OpenKeyEx(hive, key, access=flags)\n"
		"core_path = winreg.QueryValue(natlink_key, \"\")\n"
		"sys.path.append(core_path)\n"
		"winreg.CloseKey(natlink_key)\n";

	pModule = executePyCodeAsModule( codeString, moduleName );
	if( !pModule )
	{
		errorMessage = parsePyErrString();
		if( errorMessage )
		{
			m_pDragCode->displayText(
				"An exception occurred during addCoreToSysPath():\r\n" );
			m_pDragCode->displayText( errorMessage, TRUE );
			m_pDragCode->displayText( "\r\n" );
		}
		PyErr_Clear();
	}
	else
	{
		pSysModules = PySys_GetObject( "modules" );
		if( pSysModules )
		{
			PyDict_DelItemString( pSysModules, moduleName );
			Py_DECREF( pSysModules );
		}
	}
	Py_XDECREF( pModule );
	return pModule != 0;
}
