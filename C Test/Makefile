CC ?= gcc
PKGS = gtk+-3.0 libxml-2.0 libzip
CFLAGS = -Wall -Wextra -std=c11 -g $(shell pkg-config --cflags $(PKGS))
LDFLAGS = $(shell pkg-config --libs $(PKGS))

SRC = src/main.c src/docx.c
OBJ = $(SRC:.c=.o)
BIN = docxfill

.PHONY: all clean run

all: $(BIN)

$(BIN): $(OBJ)
	$(CC) $(OBJ) -o $(BIN) $(LDFLAGS)

%.o: %.c
	$(CC) $(CFLAGS) -c $< -o $@

run: $(BIN)
	./$(BIN)

clean:
	rm -f $(OBJ) $(BIN)
