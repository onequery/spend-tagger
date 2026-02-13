on run
  tell application "Terminal"
    do script "cd \"/Users/heesukim/Downloads/거래내역\"; /Users/heesukim/anaconda3/bin/python spending_tagger_app.py"
    activate
  end tell
end run
