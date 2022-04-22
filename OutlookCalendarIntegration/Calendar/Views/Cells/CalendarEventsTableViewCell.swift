//
//  CalendarEventsTableViewCell.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 09/10/2021.
//

import UIKit

class CalendarEventsTableViewCell: UITableViewCell {

    //MARK:- Outlets
    @IBOutlet weak var subjectLbl: UILabel!
    
    @IBOutlet weak var organizerLbl: UILabel!
    
    @IBOutlet weak var dateTimeLbl: UILabel!
    
    @IBOutlet weak var descriptionLbl: UILabel!
    
    override func awakeFromNib() {
        super.awakeFromNib()
        // Initialization code
    }

    override func setSelected(_ selected: Bool, animated: Bool) {
        super.setSelected(selected, animated: animated)

        // Configure the view for the selected state
    }
    
    func configCell(with data: CalendarModel) {
        self.subjectLbl.text     = "Subject : \(data.subject ?? "")"
        self.organizerLbl.text   = "Organizer : \(data.organizer ?? "")"
        self.dateTimeLbl.text    = "Start Time : \(data.startDateTime?.convertToDate() ?? "")" + "\nEnd Time   : \(data.endDatetime?.convertToDate() ?? "")"
        self.descriptionLbl.isHidden = !((data.body ?? "").count > 0)
        self.descriptionLbl.text = "Description : \(data.body ?? "")"
    }
}
