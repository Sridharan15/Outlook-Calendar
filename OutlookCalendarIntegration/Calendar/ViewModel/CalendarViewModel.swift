//
//  CalendarHelper.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 09/10/2021.
//

import Foundation
import UIKit
import MSAL
import MSGraphClientSDK
import MSGraphClientModels

class CalendarViewModel: NSObject {
    
    static var token : String = ""
    typealias getTokenCompletionBlock     = ([String : AnyObject], Error?) -> Void
    typealias createEventsCompletionBlock = (HTTPURLResponse, Error?) -> Void
    
    var calendarEvents  : [CalendarModel]!
    var reloadTableView : (() -> Void)?
    static var nextDeltaLink: String = ""
    
    override init() {
        super.init()
    }
    
    func getCalendarEvents(startingAt: String, endingAt: String, completionBlock: @escaping getTokenCompletionBlock) {
        
        //set calendar view start and end paramters
        //let startAndEndDateTime = "startDateTime=\(startingAt)&endDateTime=\(endingAt)"
        
        //MARK:-Parameters
        ///MSGraphBaseURL                - "https://graph.microsoft.com/v1.0"
        ///startAndEndDateTime           - Starting and Ending Date Time
        ///return only these values      - $select=subject,organizer,start,end
        ///sort results by date and time - $orderby=start/dateTime
        ///$top=25                       - top 25 results
        //let eventsUrlString = "\(MSGraphBaseURL)/me/events?$select=subject,organizer,start,end&$orderby=start/dateTime&$top=25"
        let eventsUrlString = CalendarViewModel.nextDeltaLink == "" ?  "https://graph.microsoft.com/v1.0/me/calendarview/delta?startdatetime=\(startingAt)&enddatetime=\(endingAt)" : CalendarViewModel.nextDeltaLink
        
        MSAuthenticationManager.sharedInstance.getTokenSilently(completion: { (accessToken, error) in
            guard let token = accessToken,error == nil else {
                return
            }
            CalendarViewModel.token = token
        })
        
        let eventsUrl     = URL(string: eventsUrlString)
        var eventsRequest = URLRequest(url: eventsUrl!)
        eventsRequest.addValue("application/json", forHTTPHeaderField: "Content-Type")
        eventsRequest.addValue("Bearer \(CalendarViewModel.token)", forHTTPHeaderField: "Authorization")
        
        let eventsDataTask = URLSession.shared.dataTask(with: eventsRequest, completionHandler: {
            data, response, error -> Void in
            //print(response!)
            do {
                let json = try JSONSerialization.jsonObject(with: data ?? Data()) as? [String : AnyObject] ?? [:]
                completionBlock(json,error)
            } catch let error {
                print(error.localizedDescription)
            }
        })
        eventsDataTask.resume()
    }
    
    func fetchCalendarEvents(completion: @escaping (() -> Void)) {
        self.calendarEvents = self.calendarEvents == nil ? [CalendarModel]() : self.calendarEvents
        self.getCalendarEvents(startingAt: "2021-11-01T13:42:45.290Z", endingAt: Date().addingTimeInterval(30*24*60*60).getISO8601DateFormat(), completionBlock: {
            jsonData, error in
            print(jsonData)
            if let eventDetails = jsonData["value"] as? [[String : AnyObject]] {
                for events in eventDetails {
                    let event = CalendarModel()
                    if let subject       = events["subject"] as? String,
                       let organizer     = events["organizer"]!["emailAddress"] as? [String:Any],
                       let startDateTime = events["start"]!["dateTime"] as? String,
                       let endDatetime   = events["end"]!["dateTime"] as? String,
                       let eventId       = events["id"] as? String,
                       let bodyPreview   = events["bodyPreview"] as? String {
                        event.id            = eventId
                        event.subject       = subject
                        event.organizer     = organizer["name"] as? String
                        event.startDateTime = startDateTime
                        event.endDatetime   = endDatetime
                        event.body          = bodyPreview
                        
                        ///if there are any update in the existing events, removing the existing event and adding the updated one
                        self.calendarEvents.removeAll(where: { $0.id == event.id })
                        self.calendarEvents.append(event)
                    }
                    
                    ///if removed reason is deleted, delete the event from the data else append the data.
                    if let eventsRemoved = events["@removed"]?["reason"] as? String, eventsRemoved == "deleted" {
                        self.calendarEvents.removeAll(where: { $0.id == (events["id"] as? String ?? "") })
                    }
                }
            }
            
            //store new delta link
            if let eventNextLink = jsonData["@odata.nextLink"] as? String {
                CalendarViewModel.nextDeltaLink = eventNextLink
            }else if let eventDeltaLink = jsonData["@odata.deltaLink"] as? String {
                CalendarViewModel.nextDeltaLink = eventDeltaLink
            }
            completion()
        })
    }
    
    func createEvents(_ subject: String, _ startDateTime: String, _ endDateTime: String, _ attendees: String, _ body: String, completionBlock : @escaping createEventsCompletionBlock) {
        
        let startDateTimeDict : [String:Any] = ["dateTime": self.removeAtInString(startDateTime).convertToDate(with: "yyyy-MM-dd'T'HH:mm") , "timeZone":TimeZone.current.identifier]
        let endDateTimeDict   : [String:Any] = ["dateTime": self.removeAtInString(startDateTime).convertToDate(with: "yyyy-MM-dd'T'HH:mm") , "timeZone":TimeZone.current.identifier]
        
        var attendeeArray = [[String:Any]]()
        let attendees = self.getAttendees(input: attendees)
            if(attendees.count > 0) {
                attendees.forEach({ email in
                    let attendeeDict : [String:Any] = ["type": "required", "emailAddress": ["address": email]]
                    attendeeArray.append(attendeeDict)
                })
        }
        
        let bodyDict : [String:Any] = ["content": body, "contentType": "text"]
        
        //Creating a dict to represent the event as parameters
        var eventsDict          = [String:Any]()
        eventsDict["subject"]   = subject
        eventsDict["start"]     = startDateTimeDict
        eventsDict["end"]       = endDateTimeDict
        eventsDict["attendees"] = attendeeArray
        eventsDict["body"]      = bodyDict
        
        var eventData = Data()
        
        do {
            eventData = try JSONSerialization.data(withJSONObject: eventsDict, options: .prettyPrinted)
        }
        catch let error {
            print(error.localizedDescription)
        }
        
        //Request
        let eventsUrlString = "\(MSGraphBaseURL)/me/events"
        let eventsUrl       = URL(string: eventsUrlString)
        
        var eventsRequest        = URLRequest(url: eventsUrl!)
        eventsRequest.httpMethod = "POST";
        eventsRequest.httpBody   = eventData;
        eventsRequest.addValue("application/json", forHTTPHeaderField: "Content-Type")
        eventsRequest.addValue("Bearer \(CalendarViewModel.token)", forHTTPHeaderField: "Authorization")
        
        let eventsDataTask = URLSession.shared.dataTask(with: eventsRequest, completionHandler: {
            data, response, error -> Void in
            print("Response----->\(response!)")
            completionBlock(response as! HTTPURLResponse,error)
        })
        eventsDataTask.resume()
        
    }
    
    func getAttendees(input:String) -> [String] {
        var attendeeArray = [String]()
        if(input.count > 0) {
            attendeeArray = input.components(separatedBy: ";")
        }
        return attendeeArray
    }
    
    func removeAtInString(_ inputString: String) -> String {
        return inputString.replacingOccurrences(of: "at", with: "")
    }
}

extension String {
    func convertToDate(with format: String = "") -> String {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = format == "" ? "yyyy-MM-dd'T'HH:mm:ss.SSSSSSS" : "MMM d, yyyy h:mm a"
        dateFormatter.timeZone   = TimeZone(abbreviation: "UTC")
        guard let date = dateFormatter.date(from:self) else { return "" }
        if(format == "") {
            dateFormatter.dateStyle = .medium
            dateFormatter.timeStyle = .short
            dateFormatter.timeZone  = TimeZone.current
        }else {
            dateFormatter.dateFormat = format
        }
        let dateString = dateFormatter.string(from: date)
        return dateString
    }
}

extension Date {
    func getISO8601DateFormat() -> String {
        let dateFormatter = DateFormatter()
        let enUSPosixLocale = Locale(identifier: "en_US_POSIX")
        dateFormatter.locale = enUSPosixLocale
        dateFormatter.dateFormat = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"
        dateFormatter.calendar = Calendar(identifier: .gregorian)

        let iso8601String = dateFormatter.string(from: self)
        return iso8601String
        
    }
}
