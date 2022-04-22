//
//  CalendarViewController.swift
//  OutlookCalendarIntegration
//
//  Created by Sridharan T on 08/10/2021.
//

import UIKit

class CalendarViewController: UIViewController {
    
    //MARK:- Outlets
    
    @IBOutlet weak var logoutBtn: UIBarButtonItem!
    
    @IBOutlet weak var refreshBtn: UIBarButtonItem!
    
    @IBOutlet weak var addEventsBtn: UIBarButtonItem!
    
    @IBOutlet weak var calendarEventsTblView: UITableView! {
        didSet {
            self.calendarEventsTblView.dataSource = self
            self.calendarEventsTblView.delegate   = self
            self.calendarEventsTblView.tableFooterView = UIView()
        }
    }
    
    let calendarVM = CalendarViewModel()
    let spinner = UIActivityIndicatorView.init(style: .large)
    
    override func viewDidLoad() {
        super.viewDidLoad()
        spinner.frame =  CGRect(x: view.center.x, y: view.center.y, width: 0, height: 0)
        self.view.addSubview(spinner)
        spinner.startAnimating()
        // fetch calendar events
        self.fetchCalendarEvents()
    }
    
    override func viewWillAppear(_ animated: Bool) {
        super.viewWillAppear(true)
    }

    //MARK:- Actions
    
    @IBAction func logoutBtnAction(_ sender: UIBarButtonItem) {
       
        MSAuthenticationManager.sharedInstance.signOut()
        
        let storyboard = UIStoryboard(name: "Main", bundle: Bundle.main)
        let loginVC = storyboard.instantiateViewController(withIdentifier: "LoginVC")
        let rootVC  = UINavigationController(rootViewController: loginVC)
        rootVC.modalPresentationStyle = .fullScreen
        self.present(rootVC, animated: true, completion: nil)

    }
    @IBAction func refreshBtnAction(_ sender: UIBarButtonItem) {
        self.refreshBtn.isEnabled = false
        spinner.startAnimating()
        fetchCalendarEvents()
    }
    
    @IBAction func didTappedAddEvents(_ sender: UIBarButtonItem) {
        let storyboard = UIStoryboard(name: "Main", bundle: Bundle.main)
        let addEventsVC = storyboard.instantiateViewController(withIdentifier: "AddNewEventsVC") as! AddNewEventsViewController
        addEventsVC.calendarVM = self.calendarVM
        addEventsVC.reloadCalendarTableView = { [weak self] in
            self?.fetchCalendarEvents()
        }
        self.present(addEventsVC, animated: true, completion: nil)
    }
    
    func fetchCalendarEvents() {
        calendarVM.fetchCalendarEvents(completion: { [weak self] in
            DispatchQueue.main.async {
                self?.spinner.stopAnimating()
                self?.calendarEventsTblView.reloadData()
                self?.refreshBtn.isEnabled = true
            }
        })
    }
}

//MARK:- TableView DataSource and Delegates
extension CalendarViewController: UITableViewDataSource, UITableViewDelegate {
    
    func tableView(_ tableView: UITableView, numberOfRowsInSection section: Int) -> Int {
        return calendarVM.calendarEvents?.count ?? 0
    }
    
    func tableView(_ tableView: UITableView, cellForRowAt indexPath: IndexPath) -> UITableViewCell {
        let cell = tableView.dequeueReusableCell(withIdentifier: "CalendarEventsCell", for: indexPath) as! CalendarEventsTableViewCell
        cell.configCell(with: calendarVM.calendarEvents[indexPath.row])
        return cell
    }
    
    
}


