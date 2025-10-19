Attribute VB_Name = "test_All"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_full_suite()
    'useful for devs - assumes Edge, Chrome, and Firefox browsers/WebDrivers are installed
    test_ActionChains.test_action_chain
    test_ActionChains.test_action_chain_sendkeys
    
    test_Alerts.test_Alerts
    test_Alerts.test_Alerts2
    
    test_Attributes.test_element_attributes_and_properties
    test_Attributes.test_css_property
    test_Attributes.test_element_aria
    
    test_Authentication.test_BasicAuthentication

    test_Capabilities.test_invisible
    test_Capabilities.test_kiosk_printing
    test_Capabilities.test_pageLoadStrategy
    'this requires a pre-existing browser launched on port 9222
    'test_Capabilities.test_remoteDebugger
    test_Capabilities.test_set_user_agent

    test_Capabilities.test_unhandled_prompts
    test_Capabilities.test_geolocation_with_incognito
    test_Capabilities.test_incognito
    test_Capabilities.test_initialize_caps_from_file
    
    test_Cookies.test_session_cookie

    test_ExecuteCDP.test_cdp_enhanced_file_download
    test_ExecuteCDP.test_cdp_enhanced_geolocation
    test_ExecuteCDP.test_cdp_enhanced_screenshot
    test_ExecuteCDP.test_cdp_random_other_stuff
    test_ExecuteCDP.test_cdp_scripts
    
    test_ExecuteCmd.test_chrome_edge_full_screenshot
    'this requires FF installation
    'test_ExecuteCmd.test_firefox_full_screenshot
    
    test_executeScript.test_call_embedded_HTML_script
    test_executeScript.test_executeScript
    test_executeScript.test_executeScriptAsync
    
    test_FileUpDownload.test_download_resource
    test_FileUpDownload.test_file_download
    test_FileUpDownload.test_file_download2
    test_FileUpDownload.test_file_upload
    
    'these requires FF installation
    'test_Firefox.test_firefox_json_viewer_bug
    'test_Firefox.test_print
    'test_Firefox.test_file_download
    
    test_Frames.test_frames_with_embed_objects
    test_Frames.test_frames_with_frameset
    test_Frames.test_frames_with_iframes
    test_Frames.test_frames_with_nested_iframes
    
    test_geolocation.test_geolocation
    
    test_highlight.test_highlight
    test_highlight.test_highlight2
    
    test_Inputs.test_select
    test_Inputs.test_radio
    
    test_IsPresent.test_IsPresent
    test_IsPresent.test_IsPresent_wait
    
    test_logging.test_logging
    
    test_PageToMethods.test_PageToHTMLMethods
    test_PageToMethods.test_PageToJSONMethods
    test_PageToMethods.test_PageToXMLMethods
    
    test_PositionSize.test_position_size
    
    test_print.test_element_screenshot
    test_print.test_print
    test_print.test_screenshot
    test_print.test_screenshot_full
    
    test_Scroll.test_scrollIntoView
    test_Scroll.test_long_scroll
    test_Scroll.test_element_scroll
    test_Scroll.test_deep_scrollIntoView
    
    test_sendkeys.test_sendkeys
    
    test_settings.test_settings
    
    test_Shadowroots.test_shadowroot

    test_Tables.test_table
    test_Tables.test_table_to_array
    test_Tables.test_large_table_to_array
    test_Tables.test_table_to_array_formatting
    
    'these two require answering prompts
    'test_UpdateDriver.test_updateDrivers
    'test_UpdateDriver.test_updateDriversForSeleniumBasic
    
    test_Wait.test_ImplicitMaxWait
    test_Wait.test_WaitForDownload
    test_Wait.test_WaitUntilDisplayed
    test_Wait.test_WaitUntilNotDisplayed
    test_Wait.test_WaitUntilNotPresent
    
    test_WebElements.test_WebElements

    test_Windows.test_Selenium_way
    test_Windows.test_url_encoding
    test_Windows.test_windows_CloseIt
    test_Windows.test_windows_Selenium_way_with_oop_approach
    test_Windows.test_windows_state
    test_Windows.test_windows_SwitchToByTitle
    test_Windows.test_windows_SwitchToByUrl
    
    'these require local extension files to test
    'test_Extensions.test_addExtensions
    'test_Extensions.test_addExtensions2
    'test_Extensions.test_InstallAddon
    
    MsgBox "tests completed"
End Sub
