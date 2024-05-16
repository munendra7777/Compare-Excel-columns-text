on run {inputFile}
    set outputFile to inputFile & ".csv"
    
    tell application "Numbers"
        activate
        open POSIX file inputFile
        tell front document
            export to POSIX file outputFile as CSV
            close saving no
        end tell
    end tell
end run
