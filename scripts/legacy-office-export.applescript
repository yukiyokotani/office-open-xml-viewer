-- Local verification only. Never used by the shipped converter.
-- The caller must provide a disposable copy and a fresh output path.
on sameLocalFile(actualPath, expectedPath)
    -- Office can return a different symbolic-link spelling of the same path.
    -- Resolve both existing files outside the application's terminology scope.
    -- Never accept only a matching basename or silently accept an unreadable path.
    try
        return ((POSIX file actualPath) as alias) = ((POSIX file expectedPath) as alias)
    on error
        return false
    end try
end sameLocalFile

on run arguments
    if (count of arguments) is not 3 then error "expected format, input and output"
    set family to item 1 of arguments
    set sourcePath to item 2 of arguments
    set outputPath to item 3 of arguments
    with timeout of 240 seconds
        if family is "xls" then
            tell application "Microsoft Excel"
                set previousAlerts to display alerts
                set previousSecurity to automation security
                if previousSecurity is missing value then error "Excel cannot report automation security; refusing to open a document"
                set workbookRef to missing value
                try
                    set display alerts to false
                    set automation security to msoAutomationSecurityForceDisable
                    if automation security is not msoAutomationSecurityForceDisable then error "Excel did not confirm disabled macros"
                    set workbookRef to open workbook workbook file name sourcePath update links do not update links read only true add to MRU false
                    save workbook as workbookRef filename outputPath file format PDF file format add to most recently used list false
                    close workbookRef saving no
                on error messageText number errorNumber
                    if workbookRef is not missing value then
                        try
                            close workbookRef saving no
                        end try
                    end if
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                    error messageText number errorNumber
                end try
                set display alerts to previousAlerts
                set automation security to previousSecurity
            end tell
        else if family is "doc" then
            tell application "Microsoft Word"
                set previousAlerts to display alerts
                set previousSecurity to automation security
                if previousSecurity is missing value then error "Word cannot report automation security; refusing to open a document"
                set previousLinks to update links at open of settings
                set documentRef to missing value
                try
                    set display alerts to alerts none
                    set automation security to msoAutomationSecurityForceDisable
                    if automation security is not msoAutomationSecurityForceDisable then error "Word did not confirm disabled macros"
                    set update links at open of settings to false
                    set documentRef to open file name sourcePath confirm conversions false read only true add to recent files false repair false
                    save as documentRef file name outputPath file format format PDF add to recent files false
                    close documentRef saving no
                on error messageText number errorNumber
                    if documentRef is not missing value then
                        try
                            close documentRef saving no
                        end try
                    end if
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                    set update links at open of settings to previousLinks
                    error messageText number errorNumber
                end try
                set display alerts to previousAlerts
                set automation security to previousSecurity
                set update links at open of settings to previousLinks
            end tell
        else if family is "ppt" then
            tell application "Microsoft PowerPoint"
                set previousSecurity to automation security
                -- Some Office for Mac builds expose this dictionary entry but
                -- return no value. Never open input when macro safety cannot
                -- be verified, and never restore an unreadable setting.
                if previousSecurity is missing value then error "PowerPoint cannot report automation security; refusing to open a document"
                set presentationRef to missing value
                try
                    set automation security to msoAutomationSecurityForceDisable
                    if automation security is not msoAutomationSecurityForceDisable then error "PowerPoint did not confirm disabled macros"
                    open POSIX file sourcePath
                    set candidateRef to active presentation
                    if not my sameLocalFile(full name of candidateRef, sourcePath) then error "PowerPoint did not open the requested disposable copy"
                    -- Arm cleanup only after ownership is established. A failed
                    -- open or focus race must never close an unrelated document.
                    set presentationRef to candidateRef
                    save presentationRef in POSIX file outputPath as save as PDF
                    close presentationRef saving no
                on error messageText number errorNumber
                    try
                        if presentationRef is not missing value then
                            close presentationRef saving no
                        end if
                    end try
                    set automation security to previousSecurity
                    error messageText number errorNumber
                end try
                set automation security to previousSecurity
            end tell
        else
            error "unsupported Office family"
        end if
    end timeout
end run
