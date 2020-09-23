/**************************************************************************
**
** Filename:     RedirStub.cpp
** Project:      CommandOutput sample
** Author:       Mattias Sjögren (mattias@mvps.org)
**               http://www.msjogren.net/dotnet/
**
** Description:  Stub application to solve redirection issues of 16 bit 
**               applications on Windows 9x.
**
**               Based on code from MS KB article Q150956
**
**               INFO: Redirection Issues on Windows 95 MS-DOS Applications
**               http://support.microsoft.com/support/kb/articles/q150/9/56.asp
**
**************************************************************************/

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <string.h>


int main(int argc, char* argv[])
{
  PROCESS_INFORMATION pi = {0};
  STARTUPINFO         si = {0};
  char *pszCmdLine;
  int i, cch = 0;

  
  if (argc < 2)                       // there has to be at least one argument
    return 1;

  for (i=1; i < argc; i++)            // get command line length
    cch += strlen(argv[i]) + 1;

  pszCmdLine = new char[++cch];

  if (!pszCmdLine)
    return 1;

  pszCmdLine[0] = '\0';
  for (i=1; i < argc; i++) {          // join all paramters in one string
    strcat(pszCmdLine, argv[i]);
    strcat(pszCmdLine, " ");
  }


  si.cb = sizeof(si);
  si.dwFlags    = STARTF_USESTDHANDLES;
  si.hStdInput  = GetStdHandle(STD_INPUT_HANDLE);
  si.hStdOutput = GetStdHandle(STD_OUTPUT_HANDLE);
  si.hStdError  = GetStdHandle(STD_ERROR_HANDLE);

  if ( CreateProcess(NULL, pszCmdLine, NULL, NULL, TRUE, 
                     0, NULL, NULL, &si, &pi) )
  {
    WaitForSingleObject(pi.hProcess, INFINITE);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
  }


  delete[] pszCmdLine;  

  return 0;
} 
