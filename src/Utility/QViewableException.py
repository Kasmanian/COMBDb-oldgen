import QView

class QViewableException():

    def throwsViewableException(func):
        def wrap(self, *args, **kwargs):
            try:
                result = func(self, *args[1:], **kwargs)
                return result
            except TypeError:
                result = func(self, *args, **kwargs)
                return result
            except Exception as e:
                return QView.showErrorScreen(e)
        return wrap