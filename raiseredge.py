from ratelimit import limits
import requests

class RaiserEdge:

    BASE_URL = 'https://api.sky.blackbaud.com'

    # Standard Tier RateLimit
    CALLS_PER_SECOND = 10

    @limits(calls=CALLS_PER_SECOND, period=1)
    def __init__(self, access_key:str=None, oauth:str=None):

        '''
        Parameters:
        access_key (str): User Access Key
        oauth (str): User OAuth Key
        '''

        assert access_key != None
        assert oauth != None

        self.__session = requests.Session()

        self.access_key = access_key
        self.oauth = oauth

        headers = {
            'Bb-Api-Subscription-Key': f'{self.access_key}',
            'Authorization': f'{self.oauth}'
        }
        
        self.__session.headers = headers

        try:
            response = self.__session.get(self.BASE_URL + '/webhook/v1/subscriptions')
            response.raise_for_status()
        except requests.exceptions.HTTPError as error:
            raise error


    # Event
    
    @limits(calls=CALLS_PER_SECOND, period=1)
    def get_event_list(self, **kwargs):
        
        try:
            response = self.__session.get(self.BASE_URL + '/event/v1/eventlist', params=kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as error:
            print(error)


    
    
    @limits(calls=CALLS_PER_SECOND, period=1)
    def create_event(self, **kwargs):
        
        assert kwargs['name'] != None
        assert kwargs['start_date'] != None
        

        if kwargs["start_time"] == "None":
            kwargs["start_time"] = ''

        try:
            response = self.__session.post(self.BASE_URL + '/event/v1/events', json=kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as error:
            print(error)

        return response.json()
    

    @limits(calls=CALLS_PER_SECOND, period=1)
    def create_participant(self, **kwargs):
        
        assert kwargs['event_id'] != None
        


        try:
            response = self.__session.post(self.BASE_URL + f'/event/v1/{kwargs["event_id"]}/participants', json=kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as error:
            print(error)

        return response.json()
    
    
    # Constituent


    @limits(calls=CALLS_PER_SECOND, period=1)
    def search_constituent(self, **kwargs):
        assert kwargs['search_text'] != None
        
        try:
            response = self.__session.get(self.BASE_URL + '/constituent/v1/constituents/search', params=kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as error:
            print(error)
        
if __name__ == "__main__":
    pass