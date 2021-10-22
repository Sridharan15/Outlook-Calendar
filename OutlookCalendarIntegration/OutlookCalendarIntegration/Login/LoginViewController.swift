//
//  ViewController.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 06/10/2021.
//

import UIKit

class LoginViewController: UIViewController {
    
    //MARK:-Outlets
    @IBOutlet weak var signInBtn: UIButton!
    @IBOutlet weak var loader   : UIActivityIndicatorView!

    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view.
        
    }

   //MARK:- Actions
    
    @IBAction func signInBtnAction(_ sender: UIButton) {
        
        loader.startAnimating()

        MSAuthenticationManager.sharedInstance.getTokenInteractivelyWithParentView(self, completion: { (accessToken, error) in

            DispatchQueue.main.async {
                
                self.loader.stopAnimating()

                if(error != nil || accessToken == nil) {
                    //show alert
                    let alert    = UIAlertController(title: "Error Signing", message: error.debugDescription, preferredStyle: .alert)
                    let OkAction = UIAlertAction(title: "OK", style: .default, handler: nil)
                    alert.addAction(OkAction)
                    self.present(alert, animated: true, completion: nil)
                    return
                }
                
               // if we got token, move to calendar events screen
                let storyboard = UIStoryboard(name: "Main", bundle: Bundle.main)
                let calendarVC = storyboard.instantiateViewController(withIdentifier: "CalendarVC")
                self.present(calendarVC, animated: true, completion: nil)
            }
        })
    }
}

