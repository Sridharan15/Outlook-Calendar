//
//  MSAuthenticationManager.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 06/10/2021.
//

import Foundation
import UIKit
import MSAL
import MSGraphClientSDK 

class MSAuthenticationManager: NSObject {
    
    static let sharedInstance = MSAuthenticationManager()
    
    var appId       : String?
    var userScopes  : Array<String>?
    var accessToken : String?
    var msalClient  : MSALPublicClientApplication?
    
    typealias getTokenCompletionBlock = (String?, Error?) -> Void
    
    override init() {
        super.init()
        // Get app ID and scopes from AuthSettings.plist
        let authConfigPath = Bundle.main.path(forResource: "AuthSettings", ofType: ".plist")
        let authConfigs    = NSDictionary(contentsOfFile: authConfigPath!) as? [String:Any] ?? [:]
        self.appId         = authConfigs["AppId"] as? String ?? ""
        self.userScopes    = authConfigs["Scopes"] as? [String] ?? []
        
        //Create the MSAL client
        do {
            self.msalClient = try MSALPublicClientApplication.init(clientId: self.appId ?? "")
        }
        catch let error {
            print(error.localizedDescription)
        }
    }
    
    func getAccessTokenForProviderOptions(authproviderOptions: MSAuthenticationProviderOptions, completion: @escaping (_ data: String?,_ error: Error?) -> Void) {
        
        self.getTokenSilently(completion: completion)
    }
    
    func getTokenInteractivelyWithParentView(_ parentView: UIViewController, completion: @escaping getTokenCompletionBlock) {
        
        let webParameters = MSALWebviewParameters.init(authPresentationViewController: parentView)
        let interactiveParameters = MSALInteractiveTokenParameters.init(scopes: self.userScopes ?? [], webviewParameters: webParameters)
        
        //Call acquireToken to open a browser so the user can sign in
        self.msalClient?.acquireToken(with: interactiveParameters, completionBlock: { (result,error) in
            
            //check result
            guard let token = result,error == nil else {
                var details = [String:Any]()
                details[NSDebugDescriptionErrorKey] = "No value was returned";
                completion(nil,NSError(domain: "MSAuthenticationManager", code: 0, userInfo: details))
                return
            }
            print("Got Token Interactively: \(token.accessToken)")
            completion(token.accessToken,nil)
        })
        
        
    }
    
    func getTokenSilently(completion: @escaping getTokenCompletionBlock) {
        
        var accounts: MSALAccount?
        
        //check if there is an account in the cache
        do {
            accounts = try self.msalClient?.allAccounts().first
        }
        catch let error {
            print(error.localizedDescription)
        }
        
        guard let account = accounts else {
            var details = [String:Any]()
            details[NSDebugDescriptionErrorKey] = "Could not retrieve account from cache";
            completion(nil,NSError(domain: "MSAuthenticationManager", code: 0, userInfo: details))
            return
        }
        
        let silentParameters = MSALSilentTokenParameters.init(scopes: self.userScopes ?? [], account: account)
        
        //Attempt to get token silently
        self.msalClient?.acquireTokenSilent(with: silentParameters, completionBlock: { (result,error) in
            
            //check result
            guard let token = result, error == nil else {
                var details = [String:Any]()
                details[NSDebugDescriptionErrorKey] = "No value was returned";
                completion(nil,NSError(domain: "MSAuthenticationManager", code: 0, userInfo: details))
                return
            }
            print("Got Token Silently: \(token.accessToken)")
            completion(token.accessToken,nil)
        })
        
    }

    func signOut() {
        var accounts = Array<MSALAccount>()
        
        //check if there is an account in the cache
        do {
            accounts = try self.msalClient?.allAccounts() ?? []
        }
        catch let error {
            print("Error getting accounts from cache,\(error.localizedDescription)")
            return
        }
        
        accounts.forEach({ account in
            do {
                try self.msalClient?.remove(account)
            }
            catch let error {
                print(error.localizedDescription)
            }
        })
        
    }
}

