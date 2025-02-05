from html2image import Html2Image
hti = Html2Image()
hti.browser.flags = [
        '--window-size=1400,9000', 
        '--disable-gpu'
    ]
hti.screenshot(
    html_file='Decard_Report_Script_output.html', 
    save_as='test_google.png'
)
