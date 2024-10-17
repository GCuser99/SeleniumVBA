Attribute VB_Name = "test_All"
Sub test_full_suite()
    'useful for devs - assumes Edge, Chrome, and Firefox browsers/WebDrivers are installed
    test_ActionChains.test_action_chain
    test_ActionChains.test_action_chain_sendkeys
    test_ActionChains.test_drag_and_drop
    
    test_Alerts.test_Alerts
    test_Alerts.test_Alerts2
    
    test_Attributes.test_element_attributes_and_properties
    test_Attributes.test_css_property
    test_Attributes.test_element_aria
    
    test_Authentication.test_BasicAuthentication
    test_Authentication.test_CDP_BasicAuthentication

    test_Capabilities.test_invisible
    test_Capabilities.test_kiosk_printing
    test_Capabilities.test_pageLoadStrategy
    test_Capabilities.test_remoteDebugger 'this leaves browser open
    test_Capabilities.test_set_user_agent
    test_Capabilities.test_unhandled_prompts
    test_Capabilities.test_geolocation_with_incognito
    test_Capabilities.test_incognito
    test_Capabilities.test_initialize_caps_from_file
    
    test_cookies.test_cookies
    test_cookies.test_cookies2
    test_cookies.test_cookies3
    
    test_Dropdowns.test_select
    
    test_ExecuteCDP.test_cdp_enhanced_file_download
    test_ExecuteCDP.test_cdp_enhanced_geolocation
    test_ExecuteCDP.test_cdp_enhanced_screenshot
    test_ExecuteCDP.test_cdp_random_other_stuff
    test_ExecuteCDP.test_cdp_scripts
    
    test_ExecuteCmd.test_chrome_edge_full_screenshot
    test_ExecuteCmd.test_firefox_full_screenshot
    
    test_executeScript.test_call_embedded_HTML_script
    test_executeScript.test_executeScript
    test_executeScript.test_executeScriptAsync
    
    test_FileUpDownload.test_download_resource
    test_FileUpDownload.test_file_download
    test_FileUpDownload.test_file_download2
    test_FileUpDownload.test_file_upload
    
    test_Firefox.test_firefox_json_viewer_bug
    test_Firefox.test_logging
    test_Firefox.test_print
    test_Firefox.test_file_download
    
    test_Frames.test_frames_with_embed_objects
    test_Frames.test_frames_with_frameset
    test_Frames.test_frames_with_iframes
    test_Frames.test_frames_with_nested_iframes
    
    test_geolocation.test_geolocation
    
    test_highlight.test_highlight
    test_highlight.test_highlight2
    
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
    
    test_Scroll.test_scroll_ops
    
    test_Sendkeys.test_Authentication
    test_Sendkeys.test_Sendkeys
    
    test_settings.test_settings
    
    test_Shadowroots.test_shadowroot
    test_Shadowroots.test_shadowroots_clear_browser_history
    
    test_Tables.test_table
    test_Tables.test_table_to_array
    test_Tables.test_table_to_array_large
    
    'these two require answering prompts
    'test_UpdateDriver.test_updateDrivers
    'test_UpdateDriver.test_updateDriversForSeleniumBasic
    
    test_Wait.test_ImplicitMaxWait
    test_Wait.test_WaitForDownload
    test_Wait.test_WaitUntilDisplayed
    test_Wait.test_WaitUntilDisplayed2
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
    
    'these require local extention files to test
    test_Extensions.test_addExtensions
    test_Extensions.test_addExtensions2
    test_Extensions.test_InstallAddon
    
    MsgBox "tests completed"
End Sub
