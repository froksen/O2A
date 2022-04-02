
class AulaBaseEvent():    
    def __init__(self) -> None:
        self.id = ""
        self.institution_code = ""
        self.creator_inst_profile_id = ""
        self.title = ""
        self.type = ""
        self.start_date_time = ""
        self.end_date_time = ""
        self.all_day = False
        self.private = False
        self.hide_in_own_calendar = False
        self.add_to_institution_calendar = False
        self.is_private = False



class AulaNewEvent(AulaBaseEvent):
    pass

        
