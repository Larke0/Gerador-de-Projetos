import win32com.client
import pythoncom

class Bricscad:
    """Main BricsCAD Automation object"""

    def __init__(self, create_if_not_exists=False, visible=True):
        """
        :param create_if_not_exists: if BricsCAD doesn't run, then
                                     a new instance will be created
        :param visible: new BricsCAD instance will be visible if True (default)
        """
        self._create_if_not_exists = create_if_not_exists
        self._visible = visible
        self._app = None

    @property
    def app(self):
        """Returns active BricsCAD application"""
        if self._app is None:
            try:
                self._app = win32com.client.GetActiveObject('BricscadApp.AcadApplication')
            except Exception as e:
                if not self._create_if_not_exists:
                    raise e
                self._app = win32com.client.Dispatch('BricscadApp.AcadApplication')
                self._app.Visible = self._visible
                print("Created new BricsCAD instance.")  # Add logging or print statements as needed
        return self._app

    @property
    def doc(self):
        """ Returns `ActiveDocument` of current :attr:`Application`"""
        return self.app.ActiveDocument

    @property
    def model(self):
        """ `ModelSpace` from active document """
        return self.doc.ModelSpace

    def iter_objects_fast(self, object_name_or_list=None, container=None, limit=None):
        """Shortcut for `iter_objects(dont_cast=True)`

        Shouldn't be used in normal situations
        """
        return self.iter_objects(object_name_or_list, container, limit, dont_cast=True)

    def iter_objects(self, object_name_or_list=None, block=None,
                     limit=None, dont_cast=False):
        """Iterate objects from `block`

        :param object_name_or_list: part of object type name, or list of it
        :param block: BricsCAD block, default - :class:`ActiveDocument.ActiveLayout.Block`
        :param limit: max number of objects to return, default infinite
        :param dont_cast: don't retrieve best interface for object, may speedup
                          iteration. Returned objects should be casted by caller
        """
        if block is None:
            block = self.doc.ActiveLayout.Block
        object_names = object_name_or_list
        if object_names:
            if isinstance(object_names, str):
                object_names = [object_names]
            object_names = [n.lower() for n in object_names]

        count = block.Count
        for i in range(count):
            item = block.Item(i)  # it's faster than `for item in block`
            if limit and i >= limit:
                return
            if object_names:
                object_name = item.ObjectName.lower()
                if not any(possible_name in object_name for possible_name in object_names):
                    continue
            if not dont_cast:
                item = self.best_interface(item)
            yield item

    def best_interface(self, obj):
        """ Retrieve best interface for object """
        return obj.QueryInterface(win32com.client.CLSIDToClassMap[win32com.client.GetClass(obj)])

# E
