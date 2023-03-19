class QPrefixGraph():
    
    def __init__(self, model):
        self.model = model
        self.__nodes__ = {}
        self.populate('Antibiotics')
        self.populate('Anaerobic')
        self.populate('Aerobic')
        self.populate('Growth')
        self.populate('B-Lac')
        self.populate('Susceptibility')
        
    def translate(self, cat, key, on, to):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        graph = [ 1, 2, 0 ]
        on = inmap[on]
        to = inmap[to]
        node = self.__nodes__[cat]
        while (graph[on]!=to):
            if key in node[on]: 
                key = node[on][key]
            else: 
                return None
            on = graph[on]
        return node[on][key] if key in node[on] else None
        
    def populate(self, type):
        typeList = self.model.selectPrefixes(type, 'Entry, Prefix, Word')
        typeEntry, typePrefix, typeWord = {}, {}, {}
        for x in typeList:
            typeEntry.update({x[0]:x[1]})
            typePrefix.update({x[1]:x[2]})
            typeWord.update({x[2]:x[0]})
        self.__nodes__[type] =  [typeEntry, typePrefix, typeWord]
        
    def get(self, cat, field):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        return list(self.__nodes__[cat][inmap[field]].keys())

    def exists(self, field, item):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        for cat in self.__nodes__:
            if item in self.__nodes__[cat][inmap[field]]:
                return True
        return False
        
# example
#pg = PrefixGraph()
#pg.__nodes__['Growth'][0][0] = 'AB'
#pg.__nodes__['Growth'][1]['AB'] = 'Burn'
#pg.__nodes__['Growth'][2]['Burn'] = 0
#print(pg.translate('Growth', 0, 'entry', 'word'))