//
//  AddNewEventsViewController.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 09/10/2021.
//

import UIKit

class AddNewEventsViewController: UIViewController {

    //MARK:- Outlets
    @IBOutlet weak var subjectTextField  : UITextField! {
        didSet {
            subjectTextField.delegate = self
        }
    }
    @IBOutlet weak var attendeesTextField: UITextField! {
        didSet {
            attendeesTextField.delegate = self
        }
    }
    @IBOutlet weak var startDateTextField: UITextField! {
        didSet {
            startDateTextField.delegate = self
            startDateTextField.placeholder = "Start Date and Time"
        }
    }
    @IBOutlet weak var startTimeTextField: UITextField! {
        didSet {
            startTimeTextField.delegate = self
            startTimeTextField.isHidden = true
        }
    }
    @IBOutlet weak var endDateTextField  : UITextField! {
        didSet {
            endDateTextField.delegate = self
            endDateTextField.placeholder = "End Date and Time"
        }
    }
    @IBOutlet weak var endTimeTextField  : UITextField! {
        didSet {
            endTimeTextField.delegate = self
            endTimeTextField.isHidden = true
        }
    }
    @IBOutlet weak var additionalInfoTextView: UITextView! {
        didSet {
            additionalInfoTextView.delegate = self
        }
    }
    
    var calendarVM     : CalendarViewModel!
    let datePicker              = DatePicker()
    var keyboardHeight : CGRect = CGRect(x: 0, y: 0, width: 0, height: 0)
    let spinner                 = UIActivityIndicatorView.init(style: .large)
    var reloadCalendarTableView : (() -> Void)?
    
    override func viewDidLoad() {
        super.viewDidLoad()
        spinner.frame =  CGRect(x: view.center.x, y: view.center.y, width: 0, height: 0)
        self.view.addSubview(spinner)
        NotificationCenter.default.addObserver(self, selector: #selector(keyboardWillChangeFrame), name: UIResponder.keyboardWillChangeFrameNotification, object: nil)
    }
    
    override func viewWillDisappear(_ animated: Bool) {
        super.viewWillDisappear(true)
        NotificationCenter.default.removeObserver(self)
    }
    
    //MARK:- Actions
    @IBAction func cancelBtnAction(_ sender: UIBarButtonItem) {
        self.dismiss(animated: true, completion: nil)
    }
    
    @IBAction func createNewEventBtnAction(_ sender: UIBarButtonItem) {
        
        spinner.startAnimating()
        
        calendarVM.createEvents(self.subjectTextField.text ?? "",
                                self.startDateTextField.text ?? "",
                                self.endDateTextField.text ?? "",
                                self.attendeesTextField.text ?? "",
                                self.additionalInfoTextView.text ?? "",
                                completionBlock: { response, error in
                                    if(error != nil) {
                                        DispatchQueue.main.async { [weak self] in
                                            self?.spinner.stopAnimating()
                                            self?.showAlert(with: error.debugDescription)
                                        }
                                    }else {
                                        DispatchQueue.main.async { [weak self] in
                                            self?.reloadCalendarTableView!()
                                            self?.dismiss(animated: true, completion: nil)
                                    }
                                }
        })
    }
    
    @objc func keyboardWillChangeFrame(notification: Notification) {
        keyboardHeight = (notification.userInfo?[UIResponder.keyboardFrameEndUserInfoKey] as? NSValue)!.cgRectValue
    }
    
    func showAlert(with message: String) {
        let alert    = UIAlertController(title: "Error Signing", message: message, preferredStyle: .alert)
        let OkAction = UIAlertAction(title: "OK", style: .default, handler: nil)
        alert.addAction(OkAction)
        self.present(alert, animated: true, completion: nil)
    }
}

//MARK:- TextField and TextView Delegates
extension AddNewEventsViewController : UITextFieldDelegate, UITextViewDelegate {
    
    func textFieldDidBeginEditing(_ textField: UITextField) {
        switch textField {
            case startDateTextField, endDateTextField:
                datePicker.createDateTimePicker(self, textField)
                datePicker.selectedDate = { date in
                    textField.text = date
                    self.hideKeyboad()
                }
                datePicker.cancelBtnTapped = {
                    self.hideKeyboad()
                }
            default:
                break
        }
    }
    
    func textField(_ textField: UITextField, shouldChangeCharactersIn range: NSRange, replacementString string: String) -> Bool {
        return true
    }
    
    func textFieldShouldReturn(_ textField: UITextField) -> Bool {
        hideKeyboad()
        return true
    }
    
    func textViewDidBeginEditing(_ textView: UITextView) {
        self.showKeyboard()
    }
    
    override func touchesBegan(_ touches: Set<UITouch>, with event: UIEvent?) {
        hideKeyboad()
    }
    
}

//MARK:- Keyboard show and dismiss methods
extension AddNewEventsViewController {
    func showKeyboard() {
        if self.view.frame.origin.y == 0 {
            self.view.frame.origin.y -= keyboardHeight.height
        }
    }
    
    func hideKeyboad() {
        if self.view.frame.origin.y != 0 {
            self.view.frame.origin.y = 0
        }
        self.view.endEditing(true)
    }
}

