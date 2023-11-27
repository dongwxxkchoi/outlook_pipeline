from langchain.document_loaders.unstructured import UnstructuredFileLoader
import os
from typing import Any, List

from langchain.docstore.document import Document


class EMLLoader(UnstructuredFileLoader):
    """Load email files using `Unstructured`.

    Works with both
    .eml and .msg files. You can process attachments in addition to the
    e-mail message itself by passing process_attachments=True into the
    constructor for the loader. By default, attachments will be processed
    with the unstructured partition function. If you already know the document
    types of the attachments, you can specify another partitioning function
    with the attachment partitioner kwarg.

    Example
    -------
    from langchain.document_loaders import UnstructuredEmailLoader

    loader = UnstructuredEmailLoader("example_data/fake-email.eml", mode="elements")
    loader.load()

    Example
    -------
    from langchain.document_loaders import UnstructuredEmailLoader

    loader = UnstructuredEmailLoader(
        "example_data/fake-email-attachment.eml",
        mode="elements",
        process_attachments=True,
    )
    loader.load()
    """

    def __init__(
        self, file_path: str, mode: str = "single", **unstructured_kwargs: Any
    ):
        process_attachments = unstructured_kwargs.get("process_attachments")
        attachment_partitioner = unstructured_kwargs.get("attachment_partitioner")

        if process_attachments and attachment_partitioner is None:
            from unstructured.partition.auto import partition

            unstructured_kwargs["attachment_partitioner"] = partition

        super().__init__(file_path=file_path, mode=mode, **unstructured_kwargs)

    def _get_elements(self) -> List:
        from unstructured.file_utils.filetype import FileType, detect_filetype

        filetype = detect_filetype(self.file_path)

        if filetype == FileType.EML:
            from unstructured.partition.email import partition_email

            return partition_email(filename=self.file_path, **self.unstructured_kwargs)
        else:
            raise ValueError(
                f"Filetype {filetype} is not supported in UnstructuredEmailLoader."
            )




