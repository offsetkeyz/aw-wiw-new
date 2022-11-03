



class time_off_request:

        def __init__(self, to_id=0, account_id=0, user_id=0, status=0, type=0, type_id=0, hours=0, start_time=0, end_time=0, cancelled_by=0, user_status=0, type_label='') -> None:
            self.to_id=to_id
            self.accoun_id=account_id
            self.user_id=user_id
            self.status=status
            self.type=type
            self.type_id=type_id
            self.hours=hours
            self.start_time=start_time
            self.end_time=end_time
            self.cancelled_by=cancelled_by
            self.user_status=user_status
            self.type_label=type_label


class shift:

    published = False
    acknowledged = False

    is_open = False
    

    def __init__(self, shift_id, account_id, user_id, location_id, position_id, site_id, start_time, end_time, published, acknowledged, notes, color, is_open) -> None:
        self.shift_id = shift_id
        self.account_id=account_id
        self.user_id=user_id
        self.location_id=location_id
        self.position_id=position_id
        self.site_id=site_id
        self.start_time=start_time
        self.end_time=end_time
        self.published=published
        self.acknowledged=acknowledged
        self.notes=notes
        self.color=color
        self.acknowledged=acknowledged
        self.is_open=is_open

        length_in_seconds = end_time - start_time
        self.length = length_in_seconds.seconds / 3600

# custom object for each employee with attributes for night hours, weekend hours, and meals          
class employee:

    def __init__(self, name, night_hours, weekend_hours, meals, overtime, row, shifts_worked, stat_hours_in):
        self.name = name
        # self.email = email
        self.night_hours = night_hours
        self.weekend_hours = weekend_hours
        self.meals = meals
        self.overtime = overtime
        self.row = row
        self.shifts_worked = shifts_worked
        self.stat_hours = stat_hours_in

class user:

        wiw_employee_id = 0
        role = 3 #only add to schedule if role = 3
        locations = [] #schedules
        is_hidden = False #only add to schedule if False
        is_active = True
        

        def __init__(self, first_name, last_name, email, wiw_employee_id, positions, role, locations, is_hidden, is_active) -> None:
            self.first_name=first_name
            self.last_name=last_name
            self.email=email
            self.positions=positions
            self.role = role
            self.locations = locations
            self.is_hidden = is_hidden
            self.is_active = is_active
            self.full_name = str(first_name) + ' ' + str(last_name)

            if wiw_employee_id == '':
                    self.wiw_employee_id = 0
            else:
                self.wiw_employee_id=int(wiw_employee_id)

        def __repr__(self) -> str:
                all_positions = {'Triage Sec Eng 1': 10470912, 'Triage Sec Analyst': 10470912, 'Triage Sec Eng 2': 10471919, 'Triage Sec Eng 3': 10474041, 'Network Ops Supp Analyst': 10477571, 'Manager, iSOC': 10477572, 'TSE4': 10486791, 'ISOC Intern': 10652403, 'EMEA Intern': 10652404, 'Triage Sec Eng 4': 10486791, 'Triage Sec Engineer 1' : 10470912, 'USA': 10654095, 'CAN': 10654096, 'DEU' : 10665016, 'GBR' : 10665016, 'Dir Business Apps Sr': 10477570, 'Co-op/ Intern':10652403, 'Tech Lead Security Svs':10474045, 'Shift Lead Security Oper':10660927, 'Team Lead Security Ops':10474045,'Concierge Sec Eng 2':10668570, 'Mgr Security Ops Sr.':10668568, 'Team Lead Tech Ops':10477572,'Mgr Security Operations':10477572, 'Mgr Concierge Services':10477572,'Mgr, Security Operations':10477572, 'Triage Business Analyst':10668571,'Concierge Sec Eng 3':10665015,'Business Sys Mgr':10477572,'Dir Security Oper Sr':10477570,'Dir Security Svs':10477570}

                output = str(self.first_name) + ' ' + str(self.last_name) + ': ' + str(self.wiw_employee_id)
                for i in self.position:
                        try:
                                output = str(output + " " + all_positions.get(i))
                        except:
                                break
                return output
