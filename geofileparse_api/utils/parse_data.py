class DocumentSplitter:
    def __init__(self, file: bytes, file_type: str):
        self.file = file
        self.file_type = file_type

    def split(self):
        if self.file_type == 'pdf':
            return self.split_pdf()
        elif self.file_type == 'doc':
            return self.split_doc()
        else:
            raise ValueError("Unsupported file type")

    def split_pdf(self):
        # PDF分割逻辑
        pass

    def split_doc(self):
        # DOC分割逻辑
        pass