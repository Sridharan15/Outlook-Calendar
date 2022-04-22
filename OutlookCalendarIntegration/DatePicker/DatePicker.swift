//
//  DatePicker.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 17/10/2021.
//

import Foundation
import UIKit

class DatePicker: NSObject {
    
    let datePicker      = UIDatePicker()
    var selectedDate    : ((String) -> Void)?
    var selectedTime    : ((String) -> Void)?
    var cancelBtnTapped : (() -> Void)?
    
    func createDateTimePicker(_ parentView:UIViewController,_ sender: UITextField) {
        
        let toolBar = UIToolbar()
        toolBar.sizeToFit()
        
        sender.inputAccessoryView = toolBar
        sender.inputView = datePicker
        datePicker.preferredDatePickerStyle = .wheels
        datePicker.minimumDate = Date()
        datePicker.datePickerMode = .dateAndTime
        
        let doneBtn = UIBarButtonItem(barButtonSystemItem: .done, target: self, action: #selector(donePressed(button:)))
        doneBtn.tag = sender.tag
        let flexibleSpace = UIBarButtonItem(barButtonSystemItem: .flexibleSpace, target: self, action: nil)
        let cancelBtn = UIBarButtonItem(barButtonSystemItem: .cancel, target: self, action: #selector(cancelPressed))
        toolBar.setItems([cancelBtn,flexibleSpace,doneBtn], animated: true)
    }
    
    @objc func donePressed(button: UIBarButtonItem) {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        formatter.timeStyle = .short
        selectedDate!(formatter.string(from: datePicker.date))
    }
        
    @objc func cancelPressed() {
        cancelBtnTapped!()
    }
}
