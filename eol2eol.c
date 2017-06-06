/*
 * Copyright (c) 2017 Eric A. Welsh. All Rights Reserved.
 *
 * eol2eol is distributed under the following BSD-style license:
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice,
 *    this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 *    this list of conditions and the following disclaimer in the documentation
 *    and/or other materials provided with the distribution.
 *
 * 3. The name of the author may not be used to endorse or promote products
 *    derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
 * IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
 * OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
 * IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
 * SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
 * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
 * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
 * OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


/*
 * Concatenates input text files together
 * Assumes there are no embedded newlines within each input line
 * If there ARE embedded newlines, they will start a new line in the output
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define TYPE_MSDOS 0x01
#define TYPE_UNIX  0x02
#define TYPE_MAC   0x04

#define MEM_OVERHEAD 1.01    /* speed hack -- overallocate to avoid reallocs */

/* realloc input string, store new max array length (including NULL) */
/* handles \r\n \n \r, including mixes of EOL characters within same file */
/* strips EOL from end of string */
char * fgets_strip_realloc(char **return_string, int *return_max_length,
                           FILE *infile)
{
    char c;
    char *string      = *return_string;
    int length        = 0;
    int total_length;
    int max_length    = *return_max_length;
    char old_c        = '\0';
    int anything_flag = 0;

    while((c = fgetc(infile)) != EOF)
    {
        anything_flag = 1;
    
        /* EOL: \n or \r\n */
        if (c == '\n')
        {
            /* MSDOS, get rid of the previously stored \r */
            if (old_c == '\r')
            {
                string[length - 1] = '\0';
            }

            old_c = c;
            
            break;
        }
        /* EOL: \r */
        /* may be a Mac text line, back up a character */
        else if (old_c == '\r')
        {
            fseek(infile, -1 * sizeof(char), SEEK_CUR);

            break;
        }
        
        old_c = c;
    
        total_length = length + 2;    /* increment, plus terminal null */

        if (total_length > max_length)
        {
            max_length = MEM_OVERHEAD * total_length;
            string     = (char *) realloc(string, max_length * sizeof(char));
        }

        string[length++] = c;
    }
    
    /* check for dangling \r from reading in Mac lines */
    if (length && string[length-1] == '\r')
    {
        string[length-1] = '\0';
    }
    
    if (length == 0)
    {
        if (!anything_flag)
        {
            return NULL;
        }
        
        if (1 > max_length)
        {
            max_length = 1;
            string     = (char *) realloc(string, sizeof(char));
        }
    }

    string[length] = '\0';
    
    *return_string = string;
    *return_max_length = max_length;

    return string;
}

int main(int argc, char *argv[])
{
    FILE *infile          = NULL;
    char *string          = NULL;
    int max_string_len    = 0;
    int eol_type;
    int i, file_index;
    int num_skipped_files = 0;
    
    int num_files         = 0;
    char **file_list      = NULL;
    
    eol_type = TYPE_MSDOS;
    for (i = 1; i < argc; i++)
    {
        if (strcmp(argv[i], "--msdos") == 0 ||
            strcmp(argv[i], "--dos") == 0)
        {
            eol_type = TYPE_MSDOS;
        }
        else if (strcmp(argv[i], "--mac") == 0 ||
                 strcmp(argv[i], "--osx") == 0)
        {
            eol_type = TYPE_MAC;
        }
        else if (strcmp(argv[i], "--unix") == 0 ||
                 strcmp(argv[i], "--posix") == 0 ||
                 strcmp(argv[i], "--linux") == 0)
        {
            eol_type = TYPE_UNIX;
        }
        else if (strcmp(argv[i], "-h") == 0 ||
                 strcmp(argv[i], "--help") == 0)
        {
            printf("Usage: eol2eol [OPTION] [FILE]...\r\n");
            printf("Concatenate EOL-converted FILE(s), or standard input, to standard output\r\n");
            printf("  Assumes there are no embedded EOL within each input line\r\n");
            printf("  If there *are* embedded EOL, they will cause a new line in the output\r\n");
            printf("\r\n");
            printf("Supported flags:\r\n");
            printf("  --dos, --msdos              convert EOL to \\r\\n (default)\r\n");
            printf("  --mac, --osx                convert EOL to \\r\r\n");
            printf("  --unix, --posix, --linux    convert EOL to \\n\r\n");
            printf("\r\n");
            printf("  -h, --help                  display this help and exit\r\n");
            printf("\r\n");
            printf("With no FILE, or when FILE is -, read standard input.\r\n");
            printf("\r\n");
            printf("Report bugs to <Eric.Welsh@moffitt.org>.\r\n");

            exit(-1);
        }
        else if (strncmp(argv[i], "--", 2) == 0)
        {
            fprintf(stderr, "eol2eol: unrecognized option `%s'\r\n", argv[i]);
            fprintf(stderr, "Try `eol2eol --help' for more information\r\n");

            exit(-1);
        }
        /* treat as a filename */
        else
        {
            file_list = (char **) realloc(file_list,
                                          (num_files + 1) * sizeof(char *));
            file_list[num_files] = argv[i];
            num_files++;
        }
    }

    /* cat all the converted files together */
    for (file_index = 0; file_index == 0 || file_index < num_files;
         file_index++)
    {
        if (file_index == 0 &&
            (num_files == 0 ||
             (num_files == 1 && strcmp(file_list[0], "-") == 0)))
        {
            infile = stdin;
        }
        else
        {
            infile = fopen(file_list[file_index], "rb");

            if (!infile)
            {
                fprintf(stderr, "eol2eol: %s: No such file or directory\r\n",
                        file_list[file_index]);
                num_skipped_files++;

                continue;
            }
        }

        switch (eol_type)
        {
            case TYPE_MSDOS:
                while(fgets_strip_realloc(&string, &max_string_len, infile))
                    printf("%s\r\n", string);
                break;

            case TYPE_UNIX:
                while(fgets_strip_realloc(&string, &max_string_len, infile))
                    printf("%s\n", string);
                break;

            case TYPE_MAC:
                while(fgets_strip_realloc(&string, &max_string_len, infile))
                    printf("%s\r", string);
                break;
        }
        
        /* Force flushing of output buffer, to make sure that Windows
         * doesn't buffer the entire file, since Windows is stupid
         */
        fflush(stdout);
        
        fclose(infile);
    }

    /* clean up cleanly, in case someone runs valgrind on this */
    if (file_list)
        free(file_list);
    if (string)
        free(string);
    
    if (num_skipped_files)
        return num_skipped_files;
    
    return 0;
}
