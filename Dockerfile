FROM python:2.7-slim
LABEL maintainer="M.B"

RUN apt-get update && apt-get install -y \
  python-poppler \
  antiword \
  catdoc \
  git \
  pkg-config \
  libpoppler-private-dev \
  libpoppler-cpp-dev \
  python-pip && pip install --upgrade setuptools cython

# Install pdf parser
RUN pip install git+https://github.com/izderadicka/pdfparser

# add utils to PATH
ENV PYTHONPATH="/document_parser/:${PYTHONPATH}"

# Place app in container.
COPY . /document_parser

# install requirements
RUN pip install -r /document_parser/requirements.txt
