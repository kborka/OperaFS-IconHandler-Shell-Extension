#pragma once
#include <Windows.h>

/// <Summary>
/// Register the in-process component in the registry.
/// </Summary>
/// <param name="pszModule">
/// Path of the module that contains the component.
/// </param>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Preferably a description instead of a key.
/// </param>
/// <param name="pszThreadModel">
/// The threading model of the component. In most cases it should be "Apartment".
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function creates the HKEY_CLASSES_ROOT\CLSID\{CLSID} key in the registry, as well as its InprocServer32 subkey.
/// </remarks>
HRESULT RegisterInprocServer(PCWSTR pszModule, const CLSID& clsid, PCWSTR pszFriendlyName, PCWSTR pszThreadModel);

/// <Summary>
/// Unegister the in-process component in the registry.
/// </Summary>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function deletes the HKEY_CLASSES_ROOT\CLSID\{CLSID} key in the registry.
/// </remarks>
HRESULT UnregisterInprocServer(const CLSID& clsid);

/// <Summary>
/// Register the IconHandler in the registry.
/// </Summary>
/// <param name="pszFileType">
/// The file type that this IconHandler will check against.
/// </param>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Should be an intermediate registry key, such as {ProgramName}.{FileType}.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function creates the following registry key structure:
/// HKEY_CLASSES_ROOT
///   {pszFileType}
///       (Default) = {pszFriendlyName}
///   {pszFriendlyName}
///       (Default) = Description
///        DefaultIcon
///            (Default) = %1
///        Shellex
///            IconHandler
///                (Default) = {clsid}
/// </remarks>
HRESULT RetisterShellExtIconHandler(PCWSTR pszFileType, const CLSID& clsid, PCWSTR pszFriendlyName);

/// <Summary>
/// Unregister the IconHandler from the registry.
/// </Summary>
/// <param name="pszFileType">
/// The file type this icon handler is associated with.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Should be an intermediate registry key such as {ProgramName}.{FileType}.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
HRESULT UnregisterShellExtIconHandler(PCWSTR pszFileType, PCWSTR pszFriendlyName);